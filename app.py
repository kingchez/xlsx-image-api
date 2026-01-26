from flask import Flask, request, send_file, jsonify
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.cell.cell import Cell
import requests
from io import BytesIO
import tempfile
import os

app = Flask(__name__)

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        "status": "running",
        "message": "XLSX Image Creator API",
        "endpoints": {
            "/create-xlsx": "POST - Create XLSX with embedded images"
        }
    })

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "healthy"}), 200

@app.route('/create-xlsx', methods=['POST'])
def create_xlsx():
    try:
        data = request.json
        items = data.get('items', [])
        
        if not items:
            return jsonify({"error": "No items provided"}), 400
        
        # Create workbook with write_only=False to ensure proper cell handling
        wb = Workbook(write_only=False)
        ws = wb.active
        ws.title = "Thumbnails"
        
        # Set column widths
        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 50
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 25
        ws.column_dimensions['F'].width = 35
        ws.column_dimensions['G'].width = 15
        ws.column_dimensions['H'].width = 25
        ws.column_dimensions['I'].width = 40
        
        # Add headers row
        headers = ['ID', 'Category', 'Title Raw', 'Title 1', 'Title 2', 'Title 3', 'Design Type', 'Thumbnail', 'Name Base64']
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            # Force the cell to use 's' type (shared string) instead of inlineStr
            # by explicitly setting data_type
            cell.data_type = 's'
        
        # Add data rows
        for idx, item in enumerate(items, start=2):
            ws.cell(row=idx, column=1, value=item.get('id', ''))
            ws.cell(row=idx, column=2, value=item.get('category_name', ''))
            ws.cell(row=idx, column=3, value=item.get('title_raw', ''))
            ws.cell(row=idx, column=4, value=item.get('title1', ''))
            ws.cell(row=idx, column=5, value=item.get('title2', ''))
            ws.cell(row=idx, column=6, value=item.get('title3', ''))
            ws.cell(row=idx, column=7, value=item.get('design_type', ''))
            # Column 8 (H) - Create empty cell with style
            empty_cell = ws.cell(row=idx, column=8, value=None)
            # Force no data type on empty cells
            empty_cell.data_type = 'n'  # numeric type for empty
            ws.cell(row=idx, column=9, value=item.get('name_base64', ''))
            
            # Download and add image
            thumbnail_url = item.get('thumbnail_url', '')
            if thumbnail_url:
                try:
                    img_response = requests.get(thumbnail_url, timeout=10)
                    if img_response.status_code == 200:
                        img_data = BytesIO(img_response.content)
                        img = XLImage(img_data)
                        img.width = 150
                        img.height = 100
                        cell = f'H{idx}'
                        ws.add_image(img, cell)
                        ws.row_dimensions[idx].height = 80
                except Exception as e:
                    print(f"Error adding image for row {idx}: {e}")
        
        # Save file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            wb.save(temp_file.name)
            temp_path = temp_file.name
        
        response = send_file(
            temp_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='thumbnails.xlsx'
        )
        
        @response.call_on_close
        def cleanup():
            try:
                os.unlink(temp_path)
            except:
                pass
        
        return response
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
