from flask import Flask, request, send_file, jsonify
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
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
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Thumbnails"
        
        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 50
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 25
        ws.column_dimensions['F'].width = 35
        ws.column_dimensions['G'].width = 15
        ws.column_dimensions['H'].width = 25
        ws.column_dimensions['I'].width = 40
        
        headers = ['ID', 'Category', 'Title Raw', 'Title 1', 'Title 2', 'Title 3', 'Design Type', 'Thumbnail', 'Name Base64']
        ws.append(headers)
        
        for idx, item in enumerate(items, start=2):
            ws.append([
                item.get('id', ''),
                item.get('category_name', ''),
                item.get('title_raw', ''),
                item.get('title1', ''),
                item.get('title2', ''),
                item.get('title3', ''),
                item.get('design_type', ''),
                '',
                item.get('name_base64', '')
            ])
            
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
```

Click **"Commit changes"**

---

3. **Second file - Click "Add file" → "Create new file"**

Name it: `requirements.txt`

Paste this:
```
Flask==3.0.0
openpyxl==3.1.2
requests==2.31.0
Pillow==10.1.0
gunicorn==21.2.0
