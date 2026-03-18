from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash, jsonify
from werkzeug.utils import secure_filename
import os
import base64
import json


PDF_TEXT_POINT_FACTOR = 0.90
PDF_TEXT_BASELINE_FACTOR = 0.84


def parse_color(color_value):
    if not color_value:
        return (0, 0, 0)
    color = str(color_value).strip().lower()
    if color.startswith('#') and len(color) == 7:
        try:
            r = int(color[1:3], 16) / 255.0
            g = int(color[3:5], 16) / 255.0
            b = int(color[5:7], 16) / 255.0
            return (r, g, b)
        except Exception:
            return (0, 0, 0)
    if color.startswith('rgb(') and color.endswith(')'):
        try:
            parts = color[4:-1].split(',')
            r = float(parts[0].strip()) / 255.0
            g = float(parts[1].strip()) / 255.0
            b = float(parts[2].strip()) / 255.0
            return (max(0, min(1, r)), max(0, min(1, g)), max(0, min(1, b)))
        except Exception:
            return (0, 0, 0)
    return (0, 0, 0)


def map_font_name(font_family):
    name = (font_family or '').strip().lower()
    normalized = name.replace('-', ' ').replace('_', ' ')
    is_bold = 'bold' in normalized
    is_italic = ('italic' in normalized) or ('oblique' in normalized)

    if 'times' in normalized or 'georgia' in normalized:
        if is_bold and is_italic:
            return 'times-bolditalic'
        if is_bold:
            return 'times-bold'
        if is_italic:
            return 'times-italic'
        return 'times-roman'

    if 'courier' in normalized:
        if is_bold and is_italic:
            return 'courier-boldoblique'
        if is_bold:
            return 'courier-bold'
        if is_italic:
            return 'courier-oblique'
        return 'courier'

    # Treat Arial and related sans-serif picks as Helvetica family in PDF base14 fonts.
    if is_bold and is_italic:
        return 'helvetica-boldoblique'
    if is_bold:
        return 'helvetica-bold'
    if is_italic:
        return 'helvetica-oblique'
    return 'helvetica'


def measure_pdf_text_width(fitz_module, text, fontname, fontsize):
    try:
        return float(fitz_module.get_text_length(text, fontname=fontname, fontsize=fontsize))
    except Exception:
        return float(len(text) * fontsize * 0.55)


def fit_text_with_ellipsis(fitz_module, text, max_width, fontname, fontsize):
    value = str(text or '')
    if measure_pdf_text_width(fitz_module, value, fontname, fontsize) <= max_width:
        return value
    ellipsis = '...'
    cut = value
    while cut and measure_pdf_text_width(fitz_module, cut + ellipsis, fontname, fontsize) > max_width:
        cut = cut[:-1]
    return (cut + ellipsis) if cut else ellipsis


def wrap_text_for_pdf(fitz_module, text, max_width, fontname, fontsize):
    raw = str(text or '').strip()
    if not raw:
        return []

    words = raw.split(' ')
    lines = []
    current = ''

    def split_long_token(token):
        chunks = []
        chunk = ''
        for ch in token:
            attempt = chunk + ch
            if measure_pdf_text_width(fitz_module, attempt, fontname, fontsize) <= max_width or not chunk:
                chunk = attempt
            else:
                chunks.append(chunk)
                chunk = ch
        if chunk:
            chunks.append(chunk)
        return chunks

    for word in words:
        if not word:
            continue

        if measure_pdf_text_width(fitz_module, word, fontname, fontsize) > max_width:
            if current:
                lines.append(current.rstrip())
                current = ''
            pieces = split_long_token(word)
            for i, piece in enumerate(pieces):
                if i < len(pieces) - 1:
                    lines.append(piece)
                else:
                    current = piece + ' '
            continue

        candidate = current + word + ' '
        if measure_pdf_text_width(fitz_module, candidate, fontname, fontsize) > max_width and current:
            lines.append(current.rstrip())
            current = word + ' '
        else:
            current = candidate

    if current:
        lines.append(current.rstrip())

    return lines

UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
EDITED_FOLDER = os.path.join(os.getcwd(), 'edited')
ALLOWED_EXTENSIONS = {'pdf'}

# ensure directories exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(EDITED_FOLDER):
    os.makedirs(EDITED_FOLDER)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'replace-with-a-secure-key'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # handle PDF upload
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            # redirect to edit page with filename
            return redirect(url_for('edit', filename=filename))
        else:
            flash('Allowed file types are pdf')
            return redirect(request.url)
    return render_template('index.html')


@app.route('/edit/<filename>')
def edit(filename):
    # Edit an uploaded/original PDF.
    return render_template(
        'edit.html',
        filename=filename,
        pdf_url=url_for('uploaded_file', filename=filename),
        project_url=None,
        source_pdf_name=filename
    )


@app.route('/edit-saved/<filename>')
def edit_saved(filename):
    safe_name = secure_filename(filename)
    pdf_path = os.path.join(EDITED_FOLDER, safe_name)
    if not os.path.exists(pdf_path):
        flash('Ficheiro editado não encontrado')
        return redirect(url_for('index'))

    project_filename = f"{safe_name}.edits.json"
    project_path = os.path.join(EDITED_FOLDER, project_filename)
    project_url = None
    source_pdf_name = safe_name

    if os.path.exists(project_path):
        project_url = url_for('edited_file', filename=project_filename)
        try:
            with open(project_path, 'r', encoding='utf-8') as pf:
                project_data = json.load(pf)
            source_pdf_name = secure_filename(project_data.get('sourcePdf') or safe_name)
        except Exception:
            source_pdf_name = safe_name

    return render_template(
        'edit.html',
        filename=safe_name,
        pdf_url=url_for('edited_file', filename=safe_name),
        project_url=project_url,
        source_pdf_name=source_pdf_name
    )


@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)


@app.route('/save', methods=['POST'])
def save_image():
    payload = request.json or {}
    data = payload.get('imageData')
    filename = payload.get('filename', 'edited.png')
    custom_pdf_name = payload.get('pdfFilename', '')
    source_pdf_name = secure_filename(payload.get('sourcePdf', ''))
    current_pdf_name = secure_filename(payload.get('currentPdf', ''))
    edits = payload.get('edits')
    editable_edits = payload.get('editableEdits')
    canvas_width = float(payload.get('canvasWidth') or 0)
    canvas_height = float(payload.get('canvasHeight') or 0)

    try:
        import fitz  # PyMuPDF

        def resolve_existing_pdf(*names):
            for raw_name in names:
                safe_name = secure_filename(raw_name or '')
                if not safe_name:
                    continue
                for folder in (UPLOAD_FOLDER, EDITED_FOLDER):
                    candidate = os.path.join(folder, safe_name)
                    if os.path.exists(candidate):
                        return candidate, safe_name
            return None, None

        base_pdf_path, base_pdf_name = resolve_existing_pdf(source_pdf_name, current_pdf_name)

        if base_pdf_path:
            original_pdf = base_pdf_path

            doc = fitz.open(original_pdf)
            page = doc[0]
            page_rect = page.rect
            if canvas_width <= 0 or canvas_height <= 0:
                sx = 1.0
                sy = 1.0
            else:
                sx = page_rect.width / canvas_width
                sy = page_rect.height / canvas_height

            for item in edits or []:
                item_type = item.get('type')
                x = float(item.get('x', 0)) * sx
                y = float(item.get('y', 0)) * sy
                w = float(item.get('w', 0)) * sx
                h = float(item.get('h', 0)) * sy
                rect = fitz.Rect(x, y, x + w, y + h)

                if item_type == 'image':
                    image_data = item.get('imageData', '')
                    if not image_data:
                        continue
                    _, encoded = image_data.split(',', 1)
                    img_binary = base64.b64decode(encoded)
                    page.insert_image(rect, stream=img_binary, keep_proportion=False, overlay=True)

                elif item_type == 'text':
                    bg = parse_color(item.get('bgColor', '#ffffff'))
                    page.draw_rect(rect, color=bg, fill=bg, overlay=True)

                    lines = item.get('lines') or []
                    if not lines and item.get('text'):
                        lines = [{
                            'text': item.get('text', ''),
                            'fontFamily': item.get('fontFamily', 'Arial'),
                            'fontSize': item.get('fontSize', 16),
                            'textColor': item.get('textColor', '#000000')
                        }]

                    pad_x = max(2.0, 10.0 * sx)
                    pad_y = max(2.0, 10.0 * sy)
                    line_gap = max(1.0, 4.0 * sy)
                    y_cursor = rect.y0 + pad_y
                    y_limit = rect.y1 - max(1.0, 6.0 * sy)
                    max_width = max(1.0, rect.width - (2.0 * pad_x))
                    use_single_line_autofit = bool(item.get('autoFitSingleLine')) and len(lines) == 1

                    for ln in lines:
                        text = (ln.get('text') or '').strip()
                        if not text:
                            continue
                        line_font_size = ln.get('fontSize', item.get('fontSize', 16))
                        font_size = max(6.0, float(line_font_size) * sy * PDF_TEXT_POINT_FACTOR)
                        color = parse_color(ln.get('textColor', item.get('textColor', '#000000')))
                        fontname = map_font_name(ln.get('fontFamily', 'Arial'))

                        if use_single_line_autofit:
                            fitted = fit_text_with_ellipsis(fitz, text, max_width, fontname, font_size)
                            fitted_w = measure_pdf_text_width(fitz, fitted, fontname, font_size)
                            draw_x = rect.x0 + pad_x + max(0.0, (max_width - fitted_w) / 2.0)
                            text_top = rect.y0 + max(pad_y, (rect.height - font_size) / 2.0)
                            page.insert_text(
                                fitz.Point(draw_x, text_top + (font_size * PDF_TEXT_BASELINE_FACTOR)),
                                fitted,
                                fontsize=font_size,
                                fontname=fontname,
                                color=color,
                                overlay=True
                            )
                            break

                        wrapped = wrap_text_for_pdf(fitz, text, max_width, fontname, font_size)

                        for i, wrapped_line in enumerate(wrapped):
                            if y_cursor + font_size > y_limit:
                                break

                            has_more = i < (len(wrapped) - 1)
                            next_y = y_cursor + font_size + line_gap
                            draw_value = wrapped_line
                            if has_more and (next_y + font_size > y_limit):
                                draw_value = fit_text_with_ellipsis(fitz, wrapped_line + ' ', max_width, fontname, font_size)
                                page.insert_text(
                                    fitz.Point(rect.x0 + pad_x, y_cursor + (font_size * PDF_TEXT_BASELINE_FACTOR)),
                                    draw_value,
                                    fontsize=font_size,
                                    fontname=fontname,
                                    color=color,
                                    overlay=True
                                )
                                y_cursor = y_limit + 1.0
                                break

                            page.insert_text(
                                fitz.Point(rect.x0 + pad_x, y_cursor + (font_size * PDF_TEXT_BASELINE_FACTOR)),
                                draw_value,
                                fontsize=font_size,
                                fontname=fontname,
                                color=color,
                                overlay=True
                            )
                            y_cursor += font_size + line_gap

                        if y_cursor + font_size > y_limit:
                            break

            original_base = os.path.splitext(base_pdf_name)[0]
        else:
            # backward-compatible fallback: raster overlay of full page image
            if not data:
                return jsonify({'status': 'error', 'message': 'Source PDF not found and no raster image data provided'}), 400
            _, encoded = data.split(',', 1)
            binary = base64.b64decode(encoded)
            safe = secure_filename(filename)

            png_name = safe
            png_path = os.path.join(EDITED_FOLDER, png_name)
            with open(png_path, 'wb') as f:
                f.write(binary)

            original_base = safe.replace('-edited.png', '')
            original_pdf = os.path.join(UPLOAD_FOLDER, original_base)
            if not original_pdf.lower().endswith('.pdf'):
                original_pdf += '.pdf'
            if not os.path.exists(original_pdf):
                doc = fitz.open()
                page = doc.new_page()
                page.insert_image(page.rect, stream=binary)
            else:
                doc = fitz.open(original_pdf)
                page = doc[0]
                page.insert_image(page.rect, stream=binary)

            png_name = safe

        if custom_pdf_name:
            custom_pdf_name = secure_filename(custom_pdf_name)
            if not custom_pdf_name.lower().endswith('.pdf'):
                custom_pdf_name += '.pdf'
            pdf_name = custom_pdf_name
        else:
            pdf_name = original_base + '-edited.pdf'
        pdf_path = os.path.join(EDITED_FOLDER, secure_filename(pdf_name))
        doc.save(pdf_path)
        doc.close()

        if (source_pdf_name or current_pdf_name) and isinstance(editable_edits, list):
            project_payload = {
                'sourcePdf': source_pdf_name or current_pdf_name,
                'editedPdf': secure_filename(pdf_name),
                'edits': editable_edits
            }
            project_name = f"{secure_filename(pdf_name)}.edits.json"
            project_path = os.path.join(EDITED_FOLDER, project_name)
            with open(project_path, 'w', encoding='utf-8') as pf:
                json.dump(project_payload, pf, ensure_ascii=False)
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

    return jsonify({'status': 'ok',
                    'png': url_for('edited_file', filename=png_name) if 'png_name' in locals() else None,
                    'pdf': url_for('edited_file', filename=secure_filename(pdf_name)),
                    'editable': url_for('edited_file', filename=f"{secure_filename(pdf_name)}.edits.json") if (source_pdf_name or current_pdf_name) and isinstance(editable_edits, list) else None,
                    'editUrl': url_for('edit_saved', filename=secure_filename(pdf_name))})

@app.route('/edited/<filename>')
def edited_file(filename):
    return send_from_directory(EDITED_FOLDER, filename)


if __name__ == '__main__':
    app.run(debug=True)
