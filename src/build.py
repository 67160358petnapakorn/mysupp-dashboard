"""ประกอบ index.html จาก src/template.html + src/embed.json (รันหลัง generate_data.py)"""
t=open('src/template.html',encoding='utf-8').read()
d=open('src/embed.json',encoding='utf-8').read()
open('index.html','w',encoding='utf-8').write(t.replace('/*__DATA__*/',d))
print('สร้าง index.html แล้ว')
