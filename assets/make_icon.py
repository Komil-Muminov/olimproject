# Рисует иконку программы: assets/icon.png и assets/icon.ico (запускать из корня проекта)
from PySide6.QtCore import Qt, QRectF
from PySide6.QtGui import QGuiApplication, QImage, QPainter, QColor, QLinearGradient, QFont, QPen
from PIL import Image
app = QGuiApplication([])
S = 256
img = QImage(S, S, QImage.Format_ARGB32); img.fill(Qt.transparent)
p = QPainter(img); p.setRenderHint(QPainter.Antialiasing)
g = QLinearGradient(0, 0, S, S); g.setColorAt(0, QColor("#3B82F6")); g.setColorAt(1, QColor("#1D4ED8"))
p.setBrush(g); p.setPen(Qt.NoPen); p.drawRoundedRect(QRectF(8, 8, 240, 240), 52, 52)
# лист-таблица
sheet = QRectF(44, 40, 168, 176)
p.setBrush(QColor("white")); p.drawRoundedRect(sheet, 18, 18)
# шапка
p.setBrush(QColor("#DBEAFE")); p.drawRoundedRect(QRectF(44, 40, 168, 44), 18, 18); p.drawRect(QRectF(44, 64, 168, 20))
# выделенный столбец ГИБ
p.setBrush(QColor("#22C55E")); p.drawRect(QRectF(156, 84, 56, 132 - 18)); p.drawRoundedRect(QRectF(156, 180, 56, 36), 18, 18); p.drawRect(QRectF(156, 180, 30, 36))
pen = QPen(QColor("#93C5FD")); pen.setWidth(5); p.setPen(pen)
for y in (84, 128, 172): p.drawLine(46, y, 210, y)
for x in (100, 156): p.drawLine(x, 84, x, 214)
p.setPen(QColor("#1E3A8A")); f = QFont("Segoe UI", 30, QFont.Black); p.setFont(f)
p.drawText(QRectF(44, 40, 168, 44), Qt.AlignCenter, "ГИБ")
p.end()
img.save("assets/icon.png")
Image.open("assets/icon.png").save("assets/icon.ico", sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])
print("ok")
