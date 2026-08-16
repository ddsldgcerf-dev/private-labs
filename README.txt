مشروع قاعدة بيانات المخابر الخاصة - نسخة أولية

التشغيل على Windows:
1) افتح PowerShell داخل مجلد المشروع.
2) python -m venv venv
3) venv\Scripts\activate
4) pip install -r requirements.txt
5) uvicorn main:app --reload
6) افتح: http://127.0.0.1:8000

ملاحظة:
قائمة المديريات الجهوية والولائية الموجودة في النسخة الأولية تجريبية فقط.
يجب استبدالها بالقائمة الرسمية التي سنعتمدها في المشروع.
