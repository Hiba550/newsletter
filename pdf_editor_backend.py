"""
PDF Editor Backend - Python-based PDF manipulation
Uses PyPDF2, reportlab, and PIL for real PDF editing
"""

from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
from PIL import Image
import os

class PDFEditor:
    def __init__(self, pdf_path):
        """Initialize PDF editor with a PDF file"""
        self.pdf_path = pdf_path
        self.reader = PdfReader(pdf_path)
        self.writer = PdfWriter()
        self.num_pages = len(self.reader.pages)
        
    def get_page_count(self):
        """Get total number of pages"""
        return self.num_pages
    
    def add_text(self, page_num, text, x, y, font_size=12, color="#000000", font_name="Helvetica"):
        """Add text to a specific page"""
        # Create overlay with text
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        
        # Set font and color
        can.setFont(font_name, font_size)
        can.setFillColor(HexColor(color))
        
        # Draw text
        can.drawString(x, y, text)
        can.save()
        
        # Merge with existing page
        packet.seek(0)
        overlay = PdfReader(packet)
        page = self.reader.pages[page_num]
        page.merge_page(overlay.pages[0])
        self.writer.add_page(page)
        
        return True
    
    def add_image(self, page_num, image_path, x, y, width, height):
        """Add image to a specific page"""
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        
        # Draw image
        can.drawImage(image_path, x, y, width=width, height=height)
        can.save()
        
        # Merge with existing page
        packet.seek(0)
        overlay = PdfReader(packet)
        page = self.reader.pages[page_num]
        page.merge_page(overlay.pages[0])
        self.writer.add_page(page)
        
        return True
    
    def add_rectangle(self, page_num, x, y, width, height, fill_color="#ffffff", stroke_color="#000000", stroke_width=1):
        """Add rectangle to a specific page"""
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        
        # Set colors
        can.setFillColor(HexColor(fill_color))
        can.setStrokeColor(HexColor(stroke_color))
        can.setLineWidth(stroke_width)
        
        # Draw rectangle
        can.rect(x, y, width, height, fill=1, stroke=1)
        can.save()
        
        # Merge with existing page
        packet.seek(0)
        overlay = PdfReader(packet)
        page = self.reader.pages[page_num]
        page.merge_page(overlay.pages[0])
        self.writer.add_page(page)
        
        return True
    
    def add_line(self, page_num, x1, y1, x2, y2, color="#000000", width=1):
        """Add line to a specific page"""
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        
        # Set color and width
        can.setStrokeColor(HexColor(color))
        can.setLineWidth(width)
        
        # Draw line
        can.line(x1, y1, x2, y2)
        can.save()
        
        # Merge with existing page
        packet.seek(0)
        overlay = PdfReader(packet)
        page = self.reader.pages[page_num]
        page.merge_page(overlay.pages[0])
        self.writer.add_page(page)
        
        return True
    
    def rotate_page(self, page_num, angle):
        """Rotate a page by specified angle"""
        page = self.reader.pages[page_num]
        page.rotate(angle)
        self.writer.add_page(page)
        return True
    
    def delete_page(self, page_num):
        """Delete a specific page"""
        for i in range(self.num_pages):
            if i != page_num:
                self.writer.add_page(self.reader.pages[i])
        return True
    
    def extract_text(self, page_num):
        """Extract text from a specific page"""
        page = self.reader.pages[page_num]
        return page.extract_text()
    
    def save(self, output_path):
        """Save the edited PDF"""
        with open(output_path, 'wb') as output_file:
            self.writer.write(output_file)
        return True
    
    def merge_pdfs(self, pdf_paths):
        """Merge multiple PDFs"""
        for pdf_path in pdf_paths:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                self.writer.add_page(page)
        return True


def create_pdf_from_html(html_content, output_path):
    """Create PDF from HTML content using reportlab"""
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    story = []
    styles = getSampleStyleSheet()
    
    # Parse HTML and add to PDF
    # This is a simplified version - you can enhance with html2pdf libraries
    para = Paragraph(html_content, styles['Normal'])
    story.append(para)
    
    doc.build(story)
    return True
