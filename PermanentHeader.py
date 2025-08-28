# Permanent Header (Standard with Jama Logo)
import os
from PyQt6.QtWidgets import (QLabel, QHBoxLayout, QFrame)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap

def permanent_header(title, file_name):
    # Create a horizontal layout for the header
    header_layout = QHBoxLayout()

    # Example permanent widget: a title label
    title_label = QLabel(f"<h1>{title}</h1>")
    title_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
    header_layout.addWidget(title_label)

    logo = QLabel()
    logo.setAlignment(Qt.AlignmentFlag.AlignRight)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    logo_path = os.path.join(script_dir, file_name)
    pixmap = QPixmap(logo_path)
    
    # Scale the pixmap while preserving the aspect ratio.
    # This will prevent distortion and blurriness.
    scaled_pixmap = pixmap.scaled(
            150,
            pixmap.height(), # Automatically determine the height
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )
    logo.setPixmap(scaled_pixmap)
    header_layout.addWidget(logo)
    
    # Add a separator line for visual clarity
    separator = QFrame()
    separator.setFrameShape(QFrame.Shape.HLine)
    separator.setFrameShadow(QFrame.Shadow.Sunken)

    return header_layout, separator