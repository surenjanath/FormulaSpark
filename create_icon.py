"""
Create FormulaSpark Icon
Generates a .ico file for the FormulaSpark executable
"""

from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor
from PyQt5.QtCore import Qt, QPoint
import sys
import os

def create_lightning_icon(size=256):
    """Create a lightning bolt icon with specified size"""
    # Create a pixmap with the specified size
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    
    # Set lightning bolt color (blue gradient)
    painter.setPen(QColor(102, 126, 234))  # #667eea
    painter.setBrush(QColor(102, 126, 234))
    
    # Scale points based on size
    scale = size / 32.0
    
    # Draw lightning bolt shape
    points = [
        (int(16 * scale), int(4 * scale)),   # Top point
        (int(10 * scale), int(16 * scale)),  # Left middle
        (int(14 * scale), int(16 * scale)),  # Right middle
        (int(8 * scale), int(28 * scale)),   # Bottom left
        (int(22 * scale), int(12 * scale)),  # Right point
        (int(18 * scale), int(12 * scale)),  # Left point
        (int(24 * scale), int(4 * scale))    # Top right
    ]
    
    # Draw the lightning bolt as a polygon
    polygon_points = [QPoint(x, y) for x, y in points]
    painter.drawPolygon(polygon_points)
    
    painter.end()
    
    return QIcon(pixmap)

def save_icon_as_ico(icon, filename, sizes=[16, 24, 32, 48, 64, 128, 256]):
    """Save icon in multiple sizes as .ico file"""
    # Create a list of pixmaps in different sizes
    pixmaps = []
    for size in sizes:
        pixmap = icon.pixmap(size, size)
        pixmaps.append(pixmap)
    
    # Save as ICO file
    icon.pixmap(256, 256).save(filename, "ICO")
    print(f"Icon saved as {filename}")

def main():
    """Main function to create the icon"""
    print("Creating FormulaSpark icon...")
    
    # Create QApplication (required for QPixmap)
    app = QApplication(sys.argv)
    
    # Create the icon
    icon = create_lightning_icon(256)
    
    # Save as ICO file
    icon_filename = "formulaspark.ico"
    save_icon_as_ico(icon, icon_filename)
    
    print(f"✅ Icon created successfully: {icon_filename}")
    print("You can now use this icon file when creating your executable!")

if __name__ == "__main__":
    main()
