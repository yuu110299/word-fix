# Word Fix VN

Web tool sửa các lỗi Word `.docx`: khoảng trắng thừa, khoảng trắng trước dấu câu, dấu phẩy dính chữ, ký hiệu `^+^+` dính chữ, chuẩn hóa Narrow 1,27 cm, Heading 2 Tahoma 11 cho Văn án/Giới thiệu và tên chương.

## Chạy local

```bash
pip install -r requirements.txt
python app.py
```

## Render

Runtime Python, build `pip install -r requirements.txt`, start `gunicorn app:app --bind 0.0.0.0:$PORT`. Chạy trực tiếp với `render.yaml` cũng được.

Lưu ý: Đoạn nhiều run được sửa theo quy tắc an toàn giữ định dạng; không có sửa chính tả ngữ nghĩa tự động. Chỉ hỗ trợ `.docx`.
