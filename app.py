# app.py

import os
import re
import base64
import io

import dash
from dash import dcc, html, callback, Input, Output, State, ALL, ctx
import dash_bootstrap_components as dbc
from docxtpl import DocxTemplate
from docx import Document

# ==================== CẤU HÌNH VÀ HẰNG SỐ ====================

# Định nghĩa các mẫu văn bản (giữ nguyên từ code gốc)
DOCUMENT_TYPES = {
    "Mẫu 1: Nghị quyết (cá biệt)": {"list": "list_1.docx", "template": "mau_1_nghi_quyet.docx"},
    "Mẫu 2: Quyết định (cá biệt) quy định trực tiếp (đối với Hội đồng, Ban được phép sử dụng dấu)": {"list": "list_2.docx", "template": "mau_2_quyet_dinh_truc_tiep_hd_ban.docx"},
    "Mẫu 3: Quyết định (cá biệt) quy định trực tiếp": {"list": "list_3.docx", "template": "mau_3_quyet_dinh_truc_tiep.docx"},
    "Mẫu 4: Quyết định (cá biệt) quy định gián tiếp": {"list": "list_4.docx", "template": "mau_4_quyet_dinh_gian_tiep.docx"},
    "Mẫu 4a: Quy chế, quy định (ban hành kèm theo quyết định)": {"list": "list_4a.docx", "template": "mau_4a_quy_che_quy_dinh.docx"},
    "Mẫu 4b: Văn bản khác (ban hành hoặc phê duyệt kèm theo quyết định)": {"list": "list_4b.docx", "template": "mau_4b_van_ban_khac.docx"},
    "Mẫu 5a: Công văn hành chính (Khi gửi đến 1 cơ quan, đơn vị)": {"list": "list_5a.docx", "template": "mau_5a_cong_van_hanh_chinh_mot.docx"},
    "Mẫu 5b: Công văn hành chính (Khi gửi đến nhiều cơ quan, đơn vị)": {"list": "list_5b.docx", "template": "mau_5b_cong_van_hanh_chinh_nhieu.docx"},
    "Mẫu 6: Văn bản có tên loại": {"list": "list_6.docx", "template": "mau_6_van_ban_co_ten_loai.docx"},
    "Mẫu 7: Tờ trình": {"list": "list_7.docx", "template": "mau_7_to_trinh.docx"},
    "Mẫu 8: Biên bản": {"list": "list_8.docx", "template": "mau_8_bien_ban.docx"},
    "Mẫu 9a: Giấy mời": {"list": "list_9a.docx", "template": "mau_9a_giay_moi.docx"},
    "Mẫu 9b: Giấy mời": {"list": "list_9b.docx", "template": "mau_9b_giay_moi.docx"},
    "Mẫu 10: Phụ lục văn bản hành chính": {"list": "list_10.docx", "template": "mau_10_phu_luc.docx"},
}
DEFAULT_CHOICE = "Mẫu 5a: Công văn hành chính (Khi gửi đến 1 cơ quan, đơn vị)"
DATA_DIR = "data"

# Khởi tạo ứng dụng Dash
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)
app.title = "Tạo văn bản tự động"

# Gán Flask server cho biến `server`
server = app.server  

# ==================== CÁC HÀM HỖ TRỢ (ĐÃ CHỈNH SỬA CHO DASH) ====================

# Các hàm read_fields_and_data_from_list_docx, extract_template_variables, create_sample_list_file
# là pure Python nên giữ nguyên, không cần thay đổi.

def read_fields_and_data_from_list_docx(file_path):
    try:
        doc = Document(file_path)
        fields = []
        data = {}
        if doc.tables:
            for table in doc.tables:
                for row_idx, row in enumerate(table.rows):
                    cells = [cell.text.strip() for cell in row.cells]
                    if row_idx == 0 and any(header in cells[0].lower() for header in ['tên trường', 'field', 'trường']):
                        continue
                    if len(cells) >= 3 and cells[0] and cells[1]:
                        field_name, field_label, field_value = cells[0], cells[1], cells[2]
                        clean_name = re.sub(r'[^\w]', '_', field_name.lower().replace(' ', '_')).strip('_')
                        if clean_name:
                            fields.append({
                                'name': clean_name,
                                'label': field_label,
                                'type': 'text_area' if any(w in field_label.lower() for w in ['nội dung', 'trích yếu', 'nơi nhận']) else 'text_input'
                            })
                            if field_value: data[clean_name] = field_value
        if not fields and not data:
            for p in doc.paragraphs:
                if ':' in p.text.strip():
                    key, value = map(str.strip, p.text.split(':', 1))
                    clean_key = re.sub(r'[^\w]', '_', key.lower().replace(' ', '_')).strip('_')
                    if clean_key:
                        fields.append({'name': clean_key, 'label': key, 'type': 'text_input'})
                        data[clean_key] = value
        return fields, data
    except Exception:
        return [], {}

def extract_template_variables(template_path):
    try:
        doc = Document(template_path)
        variables = set()
        text_content = "\n".join(p.text for p in doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text_content += "\n" + cell.text
        variables.update(re.findall(r'\{\{\s*(\w+)\s*\}\}', text_content))
        return list(variables)
    except Exception:
        return []

def create_dash_input_field(field):
    """Render input field cho Dash, thay thế render_input_field của Streamlit."""
    field_type = field.get('type', 'text_input').lower()
    field_name = field['name']
    field_label = field['label']
    
    # Dùng pattern-matching ID để callback có thể thu thập dữ liệu
    input_id = {'type': 'form-input', 'field_name': field_name}
    
    if field_type in ['text_area', 'textarea']:
        component = dbc.Textarea(id=input_id, placeholder=f"Nhập {field_label}...", style={"height": "100px"})
    # Thêm các loại khác nếu cần (date, number, etc.)
    # elif field_type == 'date':
    #     component = dcc.DatePickerSingle(id=input_id)
    else:
        component = dbc.Input(id=input_id, type="text", placeholder=f"Nhập {field_label}...")
        
    return dbc.Form([dbc.Label(field_label), component], className="mb-3")

# ==================== BỐ CỤC ỨNG DỤNG (APP LAYOUT) ====================

app.layout = dbc.Container([
    # Store để lưu trữ trạng thái trung gian, thay thế cho các biến của Streamlit
    dcc.Store(id='fields-store'),
    dcc.Store(id='data-store'),
    dcc.Store(id='template-vars-store'),
    dcc.Store(id='file-paths-store'),
    
    # Component download, sẽ được kích hoạt bởi callback
    dcc.Download(id="download-docx"),

    # Giao diện chính
    html.H1("📄 Tạo văn bản hành chính tự động (Dash Version)", className="my-4"),
    html.P("Ứng dụng đọc dữ liệu có sẵn và tự động điền vào mẫu trình bày theo quy định"),
    
    dbc.Row([
        # Cột Sidebar
        dbc.Col(
            dbc.Card([
                dbc.CardHeader("Chọn Mẫu trình bày"),
                dbc.CardBody([
                    dcc.Dropdown(
                        id='doc-type-dropdown',
                        options=[{'label': key, 'value': key} for key in DOCUMENT_TYPES.keys()],
                        value=DEFAULT_CHOICE,
                        clearable=False
                    ),
                    html.Hr(),
                    html.Div(id='file-status-display') # Nơi hiển thị trạng thái file
                ])
            ]),
            width=4,
        ),
        # Cột nội dung chính
        dbc.Col(
            [
                html.Div(id='main-content-display'), # Nơi hiển thị form hoặc nút bấm
                html.Div(id='notification-area', className="mt-3") # Nơi hiển thị thông báo
            ],
            width=8,
        )
    ]),
], fluid=True)


# ==================== CALLBACKS (LOGIC CỦA ỨNG DỤNG) ====================

# Callback 1: Phân tích file khi người dùng chọn một mẫu văn bản
@app.callback(
    Output('file-paths-store', 'data'),
    Output('fields-store', 'data'),
    Output('data-store', 'data'),
    Output('template-vars-store', 'data'),
    Output('file-status-display', 'children'),
    Input('doc-type-dropdown', 'value')
)
def analyze_files(selected_doc_type):
    if not selected_doc_type:
        return {}, [], {}, [], dbc.Alert("Vui lòng chọn một mẫu.", color="warning")

    selected_files = DOCUMENT_TYPES[selected_doc_type]
    list_path = os.path.join(DATA_DIR, selected_files["list"])
    template_path = os.path.join(DATA_DIR, selected_files["template"])
    
    paths_data = {'list_path': list_path, 'template_path': template_path}
    
    # Kiểm tra file
    template_exists = os.path.exists(template_path)
    list_exists = os.path.exists(list_path)
    
    status_alerts = []
    if template_exists:
        status_alerts.append(dbc.Alert(f"✅ Template: {template_path}", color="success", className="small"))
    else:
        status_alerts.append(dbc.Alert(f"❌ Không tìm thấy: {template_path}", color="danger", className="small"))
    
    if list_exists:
        status_alerts.append(dbc.Alert(f"✅ File dữ liệu: {list_path}", color="success", className="small"))
    else:
        status_alerts.append(dbc.Alert(f"❌ Không tìm thấy: {list_path}", color="danger", className="small"))

    if not template_exists or not list_exists:
        return paths_data, [], {}, [], status_alerts

    # Đọc dữ liệu nếu file tồn tại
    fields, data = read_fields_and_data_from_list_docx(list_path)
    template_vars = extract_template_variables(template_path)
    
    return paths_data, fields, data, template_vars, status_alerts

# Callback 2: Hiển thị form hoặc nút bấm dựa trên kết quả phân tích
@app.callback(
    Output('main-content-display', 'children'),
    Input('fields-store', 'data'),
    Input('data-store', 'data'),
    State('file-paths-store', 'data'),
)
def render_main_content(fields, data, paths):
    if not fields or not paths:
        return dbc.Alert("Chọn một mẫu để bắt đầu hoặc kiểm tra lại file.", color="info")

    list_path = paths.get('list_path', '')
    if not os.path.exists(list_path):
        return dbc.Alert(f"Không thể tiếp tục vì file '{list_path}' không tồn tại.", color="danger")
    
    fields_without_data = [f for f in fields if f['name'] not in data or not data[f['name']]]

    # Kịch bản 1: Đủ dữ liệu
    if not fields_without_data:
        return dbc.Card([
            dbc.CardHeader("🎉 Tất cả dữ liệu đã có sẵn!"),
            dbc.CardBody([
                html.P("Nhấn nút bên dưới để tạo văn bản ngay lập tức."),
                dbc.Button("🔧 Tạo văn bản từ dữ liệu có sẵn", id="generate-button", color="primary", n_clicks=0, className="w-100")
            ])
        ])

    # Kịch bản 2: Thiếu dữ liệu, hiển thị form
    else:
        form_inputs = [create_dash_input_field(field) for field in fields_without_data]
        return dbc.Card([
            dbc.CardHeader("✍️ Vui lòng nhập dữ liệu còn thiếu"),
            dbc.CardBody([
                *form_inputs,
                dbc.Button("🔧 Tạo văn bản", id="generate-button", color="primary", n_clicks=0, className="w-100")
            ])
        ])

# Callback 3: Xử lý việc tạo và tải văn bản
@app.callback(
    Output('download-docx', 'data'),
    Output('notification-area', 'children'),
    Input('generate-button', 'n_clicks'),
    State('file-paths-store', 'data'),
    State('template-vars-store', 'data'),
    State('data-store', 'data'), # Dữ liệu có sẵn
    State({'type': 'form-input', 'field_name': ALL}, 'value'), # Dữ liệu từ form
    State({'type': 'form-input', 'field_name': ALL}, 'id'), # ID của các input trong form
    prevent_initial_call=True
)
def generate_document(n_clicks, paths, template_vars, existing_data, form_values, form_ids):
    if n_clicks == 0:
        return dash.no_update, dash.no_update

    template_path = paths.get('template_path')
    if not template_path or not os.path.exists(template_path):
        return dash.no_update, dbc.Alert("Lỗi: Không tìm thấy file template.", color="danger")

    # Xây dựng context cuối cùng
    final_context = existing_data.copy()
    
    # Lấy dữ liệu từ form
    form_data = {id['field_name']: val for id, val in zip(form_ids, form_values)}
    final_context.update(form_data)

    # Đảm bảo tất cả biến đều có giá trị, tránh lỗi render
    for var in template_vars:
        if var not in final_context or not final_context[var]:
            final_context[var] = f"[{var}]" # Placeholder nếu trống

    try:
        doc = DocxTemplate(template_path)
        doc.render(final_context)
        
        # Lưu vào bộ nhớ đệm thay vì file tạm
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        # Tạo tên file thông minh
        so_ky_hieu = final_context.get('so_ky_hieu', '').replace('/', '_')
        file_name = f"cong_van_{so_ky_hieu}.docx" if so_ky_hieu else "van_ban_hoan_thanh.docx"
        
        download_data = dcc.send_bytes(file_stream.read(), file_name)
        notification = dbc.Alert("✅ Văn bản đã được tạo thành công! Bắt đầu tải xuống...", color="success")
        
        return download_data, notification

    except Exception as e:
        error_message = f"❌ Lỗi khi tạo văn bản: {str(e)}"
        return dash.no_update, dbc.Alert(error_message, color="danger")


# ==================== CHẠY ỨNG DỤNG ====================

if __name__ == '__main__':
    app.run_server(debug=True)