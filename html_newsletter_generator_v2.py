"""
HTML Newsletter Generator with Beautiful Styling
Generates professional newsletters as HTML with option for PDF conversion
"""

import pandas as pd
from jinja2 import Template
import os
import base64
from io import BytesIO
# Pillow for image resizing/compression
try:
    from PIL import Image
    HAS_PIL = True
except Exception:
    HAS_PIL = False

# PDF generation via browser print (no pdfkit needed)


class HTMLNewsletterGenerator:
    def __init__(self, excel_path, image_paths, session_id):
        self.excel_path = excel_path
        self.image_paths = image_paths
        self.session_id = session_id
        self.data = {}
        self._load_data()
        
    def _load_data(self):
        """Load all Excel data"""
        try:
            excel_file = pd.ExcelFile(self.excel_path)

            # Newsletter Info
            df = pd.read_excel(self.excel_path, sheet_name='Newsletter Info')
            self.data['info'] = {row['Field']: row['Value'] for _, row in df.iterrows() if pd.notna(row['Field'])}

            # Editorial Board
            df = pd.read_excel(self.excel_path, sheet_name='Editorial Board')
            self.data['editorial'] = df.dropna(subset=['Role']).to_dict('records')

            # Vision & Mission
            if 'Vision & Mission' in excel_file.sheet_names:
                df = pd.read_excel(self.excel_path, sheet_name='Vision & Mission')
                self.data['vision_mission'] = df.dropna(subset=['Type']).to_dict('records')
            else:
                self.data['vision_mission'] = []

            # Program Objectives (PEO)
            if 'Program Objectives' in excel_file.sheet_names:
                df = pd.read_excel(self.excel_path, sheet_name='Program Objectives')
                # Normalize column names: some templates use 'Objective' while templates expect 'Description'
                df = df.rename(columns=lambda c: c.strip() if isinstance(c, str) else c)
                df = df.rename(columns={'Objective': 'Description', 'Outcome': 'Description'})
                self.data['peo'] = df.dropna(subset=['Code']).to_dict('records')
            else:
                self.data['peo'] = []

            # Program Outcomes (PSO)
            if 'Program Outcomes' in excel_file.sheet_names:
                df = pd.read_excel(self.excel_path, sheet_name='Program Outcomes')
                # Normalize column names: templates expect 'Description' for display
                df = df.rename(columns=lambda c: c.strip() if isinstance(c, str) else c)
                df = df.rename(columns={'Outcome': 'Description', 'Objective': 'Description'})
                self.data['pso'] = df.dropna(subset=['Code']).to_dict('records')
            else:
                self.data['pso'] = []

            # Department Events
            df = pd.read_excel(self.excel_path, sheet_name='Department Events')
            self.data['events'] = df.dropna(subset=['Event Title']).to_dict('records')

            # Contact Info
            if 'Contact Info' in excel_file.sheet_names:
                df = pd.read_excel(self.excel_path, sheet_name='Contact Info')
                self.data['contact'] = {row['Field']: row['Value'] for _, row in df.iterrows() if pd.notna(row['Field'])}
            else:
                self.data['contact'] = {}

        except Exception as e:
            raise Exception(f"Error loading Excel: {str(e)}")
    
    def _convert_images_to_base64(self):
        """Convert images to base64 for embedding in HTML.

        Returns a tuple (embedded_images, image_paths) where:
        - embedded_images: dict mapping image keys -> base64 data (optimized/resized)
        - image_paths: dict mapping header keys -> relative file path (preferred for large header images)
        """
        embedded_images = {}
        image_paths = {}

        # Header images (prefer referencing the file path so the HTML doesn't inline large binaries)
        header_keys = {
            'college_logo': os.path.join('static', 'images', 'college_logo.png'),
            'orbits_logo': os.path.join('static', 'images', 'orbits_logo.png'),
            'naac_badge': os.path.join('static', 'images', 'naac_badge.png'),
            'vision': os.path.join('static', 'images', 'vision.png'),
        }

        for key, path in header_keys.items():
            try:
                if os.path.exists(path):
                    # Use relative path (so browsers load the image from disk) to avoid inlining huge images
                    image_paths[key] = os.path.join('static', 'images', os.path.basename(path)).replace('\\', '/')
            except Exception:
                pass

        # Helper to open and optionally resize/compress images before base64-encoding
        def _encode_image(path, max_width=1000, quality=75):
            try:
                if HAS_PIL:
                    with Image.open(path) as im:
                        im_format = 'PNG' if im.format == 'PNG' else 'JPEG'
                        # Resize if too large
                        w, h = im.size
                        if w > max_width:
                            new_h = int(max_width * h / w)
                            im = im.resize((max_width, new_h), Image.LANCZOS)

                        bio = BytesIO()
                        if im_format == 'JPEG':
                            im = im.convert('RGB')
                            im.save(bio, format='JPEG', quality=quality, optimize=True)
                        else:
                            # For PNG preserve transparency but reduce size by saving with optimize
                            im.save(bio, format='PNG', optimize=True)
                        return base64.b64encode(bio.getvalue()).decode()
                else:
                    # Fallback: raw read
                    with open(path, 'rb') as f:
                        return base64.b64encode(f.read()).decode()
            except Exception:
                return None

        # Then include any other images provided by the caller (e.g., event images, main image)
        for key, path in self.image_paths.items():
            # skip if a header image path already exists for same key (we prefer file path for headers)
            if key in image_paths:
                continue
            try:
                if os.path.exists(path):
                    # Compress/resize big images before embedding to keep HTML size reasonable
                    encoded = _encode_image(path, max_width=1000, quality=75)
                    if encoded:
                        embedded_images[key] = encoded
            except Exception:
                # ignore invalid image paths or read errors
                pass

        return embedded_images, image_paths
    
    def _group_events_by_section(self):
        """Group events by department/section"""
        sections = {}
        for event in self.data['events']:
            section = event.get('Department/Section', 'OTHER ACTIVITIES')
            if pd.notna(section):
                if section not in sections:
                    sections[section] = []
                sections[section].append(event)
        return sections
    
    def _get_vision_mission_by_type(self):
        """Separate vision/mission items by type"""
        vision = [item for item in self.data['vision_mission'] if 'vision' in str(item.get('Type', '')).lower()]
        mission_items = [item for item in self.data['vision_mission'] if 'mission' in str(item.get('Type', '')).lower()]
        return vision, mission_items
    
    def _build_event_details(self, event):
        """Build event details list"""
        details = []
        speaker = event.get('Guest Speaker')
        if pd.notna(speaker) and str(speaker).lower() != 'nan':
            details.append(f"Guest Speaker: {speaker}")
        
        location = event.get('Location')
        if pd.notna(location) and str(location).lower() != 'nan':
            details.append(f"Location: {location}")
        
        return details
    
    def _clean_repetitive_text(self, text):
        """Remove repetitive consecutive sentences from text - AGGRESSIVE VERSION"""
        if not text or pd.isna(text):
            return text
        
        text = str(text).strip()
        if not text:
            return text
        
        # More robust sentence splitting using regex
        import re
        # Split on period, exclamation, or question mark followed by space or end
        sentences = re.split(r'([.!?]+)\s+', text)
        
        # Reconstruct sentences with their punctuation
        full_sentences = []
        for i in range(0, len(sentences) - 1, 2):
            if sentences[i].strip():
                sentence = sentences[i].strip()
                if i + 1 < len(sentences):
                    sentence += sentences[i + 1]
                full_sentences.append(sentence.strip())
        
        # If last element doesn't have punctuation, add it
        if len(sentences) % 2 == 1 and sentences[-1].strip():
            full_sentences.append(sentences[-1].strip())
        
        # Remove ALL duplicates (not just consecutive)
        seen = set()
        cleaned = []
        for sentence in full_sentences:
            # Normalize: lowercase, remove extra whitespace
            normalized = ' '.join(sentence.lower().split())
            
            # Only add if we haven't seen this before
            if normalized and normalized not in seen and len(normalized) > 5:
                seen.add(normalized)
                cleaned.append(sentence)
        
        result = ' '.join(cleaned)
        
        return result if result else text
    
    
    def generate_html(self):
        """Generate complete HTML newsletter"""
        embedded_images, image_file_paths = self._convert_images_to_base64()
        sections = self._group_events_by_section()
        vision, mission = self._get_vision_mission_by_type()
        
        # Get static image data or file paths (prefer file paths for large header images)
        college_logo_b64 = embedded_images.get('college_logo', '')
        orbits_logo_b64 = embedded_images.get('orbits_logo', '')
        naac_badge_b64 = embedded_images.get('naac_badge', '')
        vision_b64 = embedded_images.get('vision', '')
        
        # Get front image from Excel "Front Image" field
        front_image_field = self.data['info'].get('Front Image', '1.png')
        # Extract the key (filename without extension)
        front_image_key = os.path.splitext(str(front_image_field))[0].lower() if front_image_field else '1'
        main_image = embedded_images.get(front_image_key, '')

        college_logo_path = image_file_paths.get('college_logo') if image_file_paths else None
        orbits_logo_path = image_file_paths.get('orbits_logo') if image_file_paths else None
        naac_badge_path = image_file_paths.get('naac_badge') if image_file_paths else None
        vision_path = image_file_paths.get('vision') if image_file_paths else None
        
        html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Newsletter</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;500;600;700;800;900&family=Inter:wght@300;400;500;600;700;800&family=Roboto+Slab:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        /* ==================== PROFESSIONAL NEWSLETTER CSS ==================== */
        /* Reset & Base Styles */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        :root {
            /* Modern Template 2 - Teal & Purple Theme */
            --primary-teal: #0d9488;
            --primary-teal-light: #14b8a6;
            --primary-teal-dark: #0f766e;
            --accent-purple: #7c3aed;
            --accent-purple-light: #8b5cf6;
            --accent-purple-dark: #6d28d9;
            --accent-coral: #f97316;
            --accent-coral-light: #fb923c;
            --text-dark: #0f172a;
            --text-medium: #475569;
            --text-light: #64748b;
            --bg-white: #ffffff;
            --bg-light: #f8fafc;
            --bg-gray: #f1f5f9;
            --border-light: #e2e8f0;
            
            /* Gradients */
            --gradient-teal: linear-gradient(135deg, #0d9488 0%, #14b8a6 100%);
            --gradient-purple: linear-gradient(135deg, #7c3aed 0%, #8b5cf6 100%);
            --gradient-sunset: linear-gradient(135deg, #f97316 0%, #fb923c 100%);
            
            /* Shadows */
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.08);
            --shadow-md: 0 4px 12px rgba(0,0,0,0.1);
            --shadow-lg: 0 10px 20px rgba(0,0,0,0.12);
            --shadow-card: 0 2px 8px rgba(0,0,0,0.06);
        }
        
        @page {
            size: A4;
            margin: 0;
        }

        @media print {
            * {
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }
            
            body {
                background: white !important;
                padding: 0 !important;
            }
            
            .a4 {
                box-shadow: none !important;
                margin: 0 !important;
                height: auto !important;
                max-height: none !important;
                min-height: 0 !important;
                overflow: visible !important;
                page-break-after: auto;
                break-after: auto;
                page-break-inside: auto; /* Allow breaks inside for very long content */
                break-inside: auto;
            }
            
            .a4.last-page {
                page-break-after: avoid !important;
                break-after: avoid !important;
            }
            
            /* Prevent orphaned section headers */
            .section-header {
                page-break-after: avoid !important;
                break-after: avoid !important;
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }
            
            .section-title {
                page-break-after: avoid !important;
                break-after: avoid !important;
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }
            
            .section-underline {
                page-break-before: avoid !important;
                break-before: avoid !important;
            }
            
            /* Keep event cards intact - no mid-card breaks */
            .event-card {
                page-break-inside: avoid !important;
                break-inside: avoid !important;
                page-break-before: auto;
                break-before: auto;
            }
            
            .event-title {
                page-break-after: avoid !important;
                break-after: avoid !important;
            }
            
            /* Keep PEO/PSO items together */
            .peo-item, .pso-item {
                page-break-inside: avoid !important;
                break-inside: avoid !important;
                page-break-before: auto;
                break-before: auto;
            }
            
            /* Mission list items */
            .mission-list li {
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }
            
            /* Editorial board members */
            .board-member {
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }
            
            /* Images should not break */
            .main-image, .event-image {
                page-break-inside: avoid !important;
                break-inside: avoid !important;
                page-break-before: auto;
                break-before: auto;
            }
            
            /* Table of contents */
            .contents-table {
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }
            
            /* Header should stay with content */
            .header {
                page-break-after: avoid !important;
                break-after: avoid !important;
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }
            
            /* Orphan and widow control for text */
            p, li {
                orphans: 3;
                widows: 3;
            }
            
            /* Manual page break utility */
            .page-break {
                display: block !important;
                height: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
                page-break-after: always !important;
                break-after: page !important;
            }
            
            /* Utility: avoid break */
            .avoid-break {
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }
            
            /* Utility: force break before */
            .break-before {
                page-break-before: always !important;
                break-before: page !important;
            }
            
            /* Utility: force break after */
            .break-after {
                page-break-after: always !important;
                break-after: page !important;
            }
        }
        
        /* Body & Container */
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            line-height: 1.7;
            color: var(--text-dark);
            font-size: 10.5pt;
            background: linear-gradient(145deg, #e8eef5 0%, #f0f4f8 50%, #e8eef5 100%);
            background-attachment: fixed;
            padding: 20px;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }

        /* A4 Page Container */
        .a4 {
            width: 100%;
            max-width: 210mm;
            min-height: 297mm;
            margin: 0 auto 24px;
            background: var(--bg-white);
            box-shadow: var(--shadow-lg), 0 0 0 1px rgba(0,0,0,0.03);
            padding: 10mm 12mm;
            position: relative;
            border-radius: 2px;
            overflow: visible; /* Allow content to flow naturally */
        }
        
        .a4.last-page {
            /* No special styling needed for screen view */
        }

        .a4::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: var(--gradient-blue);
        }

        .a4 img {
            max-width: 100%;
            height: auto;
            display: block;
            margin-left: auto;
            margin-right: auto;
        }
        
        /* ==================== COMPACT MODERN HEADER (TEMPLATE 2) ==================== */
        .header {
            background: white;
            border: none;
            border-radius: 10px;
            margin-bottom: 20px;
            page-break-inside: avoid;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            overflow: hidden;
            position: relative;
        }
        
        /* Thin top accent */
        .header::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: var(--gradient-teal);
        }
        
        .header-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 18px 24px 16px;
            background: white;
            gap: 20px;
            position: relative;
        }
        
        .header-logo {
            flex: 0 0 auto;
            text-align: center;
        }
        
        .header-logo img {
            max-height: 55px;
            width: auto;
            display: block;
        }
        
        .header-center {
            flex: 1;
            text-align: center;
        }

        .header-center img {
            max-height: 60px;
            margin: 0 auto 8px;
            display: block;
        }
        
        .header-center p {
            font-family: 'Inter', sans-serif;
            font-size: 7.5pt;
            font-weight: 600;
            color: var(--text-medium);
            margin: 3px 0;
            line-height: 1.5;
            letter-spacing: 0.8px;
            text-transform: uppercase;
        }
        
        .header-center p:first-of-type {
            font-size: 8.5pt;
            font-weight: 700;
            color: var(--text-dark);
            letter-spacing: 1px;
            margin-bottom: 6px;
        }
        
        .header-badge {
            flex: 0 0 auto;
            text-align: center;
        }
        
        .header-badge img {
            max-height: 50px;
            width: auto;
            display: block;
        }
        
        /* Compact bottom bar */
        .header-bottom {
            display: flex;
            justify-content: space-between;
            padding: 12px 24px;
            background: var(--gradient-purple);
            font-family: 'Inter', sans-serif;
            font-weight: 700;
            font-size: 9pt;
            color: white;
            letter-spacing: 1.2px;
            text-transform: uppercase;
        }
        
        .header-bottom-left { 
            text-align: left; 
            flex: 1;
        }
        
        .header-bottom-center { 
            text-align: center; 
            flex: 1; 
            font-weight: 800;
        }
        
        .header-bottom-right { 
            text-align: right; 
            flex: 1;
        }
        
        /* ==================== MODERN SECTION TITLES (TEMPLATE 2) ==================== */
        .section-header {
            background: var(--gradient-teal);
            color: white;
            padding: 16px 28px;
            margin: 32px -12mm 24px -12mm;
            text-align: center;
            position: relative;
            box-shadow: var(--shadow-sm);
            border-radius: 0;
        }
        
        .section-header h2 {
            font-family: 'Playfair Display', Georgia, serif;
            font-size: 19pt;
            font-weight: 700;
            letter-spacing: 2.5px;
            text-transform: uppercase;
            margin: 0;
            text-shadow: 0 2px 4px rgba(0,0,0,0.15);
        }
        
        .section-title {
            text-align: center;
            font-family: 'Playfair Display', Georgia, serif;
            font-size: 21pt;
            font-weight: 700;
            color: var(--primary-teal);
            margin: 32px 0 16px;
            page-break-inside: avoid;
            letter-spacing: 1.8px;
            text-transform: uppercase;
            position: relative;
            padding-bottom: 20px;
        }
        
        .section-title::after {
            content: '';
            display: block;
            width: 80px;
            height: 5px;
            background: var(--gradient-sunset);
            margin: 16px auto 0;
            border-radius: 3px;
        }
        
        .section-title.red {
            color: var(--accent-purple);
        }
        
        .section-title.red::after {
            background: var(--gradient-purple);
        }
        
        .section-underline {
            text-align: center;
            font-size: 8.5pt;
            color: var(--text-light);
            margin-bottom: 24px;
            letter-spacing: 5px;
            page-break-inside: avoid;
        }
        
        /* ==================== CONTENT BLOCKS ==================== */
        .content {
            margin: 0 0 16px 0;
            text-align: justify;
            font-size: 10.5pt;
            line-height: 1.8;
            color: var(--text-medium);
        }
        
        .content p {
            margin-bottom: 12px;
        }
        
        .mission-list {
            margin: 14px 0 18px 28px;
            padding-left: 0;
        }
        
        .mission-list li {
            list-style-type: none;
            margin-bottom: 10px;
            font-size: 10.5pt;
            color: var(--text-medium);
            line-height: 1.7;
            padding-left: 24px;
            position: relative;
        }
        
        .mission-list li::before {
            content: '▸';
            position: absolute;
            left: 0;
            color: var(--accent-gold-dark);
            font-weight: bold;
            font-size: 12pt;
        }
        
        /* ==================== ENHANCED PEO & PSO ITEMS ==================== */
        .peo-item, .pso-item {
            margin: 16px 0;
            font-size: 10.5pt;
            page-break-inside: avoid;
            padding: 16px 20px;
            background: linear-gradient(135deg, #eff6ff 0%, var(--bg-white) 100%);
            border-left: 5px solid var(--primary-blue-light);
            border-radius: 0 8px 8px 0;
            box-shadow: var(--shadow-sm);
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }
        
        .peo-item::before, .pso-item::before {
            content: '';
            position: absolute;
            top: 0;
            right: 0;
            width: 60px;
            height: 60px;
            background: linear-gradient(135deg, transparent 50%, rgba(30, 64, 175, 0.05) 50%);
        }
        
        .pso-item {
            background: linear-gradient(135deg, #fef2f2 0%, var(--bg-white) 100%);
            border-left-color: var(--accent-red);
        }
        
        .pso-item::before {
            background: linear-gradient(135deg, transparent 50%, rgba(185, 28, 28, 0.05) 50%);
        }
        
        .peo-item strong, .pso-item strong {
            font-family: 'Inter', sans-serif;
            color: var(--primary-blue);
            font-weight: 800;
            display: inline-block;
            margin-bottom: 8px;
            font-size: 11.5pt;
            letter-spacing: 0.5px;
            background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-blue-light) 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        
        .pso-item strong {
            background: linear-gradient(135deg, var(--accent-red) 0%, var(--accent-red-light) 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        
        .peo-item p, .pso-item p {
            color: var(--text-medium);
            margin: 0;
            font-size: 10pt;
            line-height: 1.7;
        }
        
        /* ==================== ENHANCED EDITORIAL BOARD ==================== */
        .editorial-board {
            margin: 20px 0;
            background: linear-gradient(135deg, var(--bg-light) 0%, var(--bg-white) 100%);
            padding: 20px;
            border-radius: 12px;
            border: 1px solid var(--border-light);
            box-shadow: var(--shadow-sm);
        }
        
        .board-member {
            margin: 12px 0;
            font-size: 10.5pt;
            line-height: 1.7;
            padding: 12px 16px;
            background: var(--bg-white);
            border-radius: 8px;
            border-left: 4px solid var(--accent-gold);
            box-shadow: var(--shadow-sm);
            transition: all 0.2s ease;
        }
        
        .board-member:hover {
            transform: translateX(4px);
            box-shadow: var(--shadow-md);
        }
        
        .board-member:last-child {
            margin-bottom: 0;
        }
        
        .board-member strong {
            font-family: 'Inter', sans-serif;
            font-weight: 700;
            color: var(--primary-blue);
            min-width: 160px;
            display: inline-block;
        }
        
        /* ==================== ENHANCED IMAGES ==================== */
        .main-image {
            text-align: center;
            margin: 24px auto;
            page-break-inside: avoid;
        }
        
        .main-image img {
            max-width: 90%;
            max-height: 260px;
            text-align: center;
            margin: 18px auto;
            page-break-inside: avoid;
        }
        
        .event-image img {
            max-width: 85%;
            max-height: 260px;
            height: auto;
            border: 3px solid var(--border-light);
            padding: 6px;
            box-shadow: var(--shadow-md);
            background: var(--bg-white);
            border-radius: 8px;
            display: block;
            margin: 0 auto;
            transition: transform 0.3s ease;
        }
        
        /* ==================== ENHANCED TABLE OF CONTENTS ==================== */
        .contents-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            margin: 24px 0;
            page-break-inside: avoid;
            box-shadow: var(--shadow-md);
            border-radius: 12px;
            overflow: hidden;
        }
        
        .contents-table th {
            background: var(--gradient-blue);
            color: white;
            padding: 16px 18px;
            text-align: center;
            font-family: 'Inter', sans-serif;
            font-size: 10pt;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .contents-table td {
            border-bottom: 1px solid var(--border-light);
            padding: 14px 18px;
            text-align: center;
            font-size: 10.5pt;
            font-weight: 600;
            color: var(--primary-blue);
            background: var(--bg-white);
            transition: background 0.2s ease;
        }

        .contents-table tbody tr:nth-child(odd) td {
            background: var(--bg-light);
        }
        
        .contents-table tbody tr:hover td {
            background: #eff6ff;
        }

        .contents-table tbody tr:last-child td {
            border-bottom: none;
        }

        .contents-table th:first-child, .contents-table td:first-child {
            text-align: center;
            font-weight: 800;
            width: 10%;
        }

        .contents-table th:nth-child(2), .contents-table td:nth-child(2) {
            text-align: left;
            padding-left: 28px;
            letter-spacing: 0.5px;
        }

        .contents-table th:nth-child(3), .contents-table td:nth-child(3) {
            text-align: center;
            font-weight: 800;
            width: 12%;
        }
        
        /* ==================== ENHANCED CONTACT INFO ==================== */
        .contact-info {
            font-size: 10.5pt;
            line-height: 1.8;
            margin-top: 24px;
            padding: 24px 28px;
            background: linear-gradient(135deg, var(--bg-cream) 0%, var(--bg-white) 100%);
            border-left: 5px solid var(--accent-gold);
            border-radius: 0 12px 12px 0;
            box-shadow: var(--shadow-sm);
            position: relative;
            overflow: hidden;
        }
        
        .contact-info::before {
            content: '';
            position: absolute;
            top: 0;
            right: 0;
            width: 100px;
            height: 100px;
            background: linear-gradient(135deg, transparent 50%, rgba(249, 168, 37, 0.1) 50%);
        }
        
        .contact-info p {
            margin: 8px 0;
            color: var(--text-medium);
        }
        
        .contact-info strong {
            color: var(--primary-blue);
            font-weight: 700;
        }
        
        /* ==================== ENHANCED FOOTER ==================== */
        .footer-board {
            display: flex;
            justify-content: space-around;
            margin-top: 48px;
            padding-top: 24px;
            border-top: 3px solid var(--accent-gold);
            page-break-inside: avoid;
            gap: 16px;
        }
        
        .footer-board-member {
            text-align: center;
            font-size: 10pt;
            flex: 1;
            padding: 16px 12px;
            background: linear-gradient(135deg, var(--bg-light) 0%, var(--bg-white) 100%);
            border-radius: 8px;
            box-shadow: var(--shadow-sm);
            transition: transform 0.2s ease;
        }
        
        .footer-board-member:hover {
            transform: translateY(-2px);
        }
        
        .footer-board-member .name {
            font-family: 'Inter', sans-serif;
            font-weight: 700;
            margin-bottom: 8px;
            color: var(--primary-blue);
            font-size: 11pt;
            letter-spacing: 0.3px;
        }
        
        .footer-board-member .role {
            font-size: 9pt;
            color: var(--text-light);
            font-style: italic;
            line-height: 1.5;
        }
        
        /* ==================== SOCIAL MEDIA ==================== */
        .social-media {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 16px;
            margin: 10px 0;
            flex-wrap: wrap;
        }
        
        .social-icon {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 32px;
            height: 32px;
            border-radius: 50%;
            background: linear-gradient(135deg, var(--primary-blue) 0%, #1a365d 100%);
            color: white;
            font-size: 14px;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
            box-shadow: var(--shadow-sm);
        }
        
        .social-icon:hover {
            transform: scale(1.1);
            box-shadow: var(--shadow-md);
        }
        
        .social-icon.youtube {
            background: linear-gradient(135deg, #FF0000 0%, #CC0000 100%);
        }
        
        .social-icon.instagram {
            background: linear-gradient(135deg, #E4405F 0%, #C13584 50%, #833AB4 100%);
        }
        
        .social-icon.linkedin {
            background: linear-gradient(135deg, #0077B5 0%, #005582 100%);
        }
        
        .social-icon.twitter {
            background: linear-gradient(135deg, #000000 0%, #333333 100%);
        }
        
        .social-handle {
            font-family: 'Inter', sans-serif;
            font-size: 10pt;
            font-weight: 600;
            color: var(--primary-blue);
            margin-top: 6px;
            text-align: center;
        }
        
        /* ==================== UTILITIES ==================== */
        hr {
            border: none;
            height: 2px;
            background: linear-gradient(to right, transparent, var(--border-light), transparent);
            margin: 20px 0;
        }
        
        .divider {
            height: 1px;
            background: linear-gradient(to right, var(--accent-gold), var(--primary-blue), var(--accent-gold));
            margin: 24px 0;
            opacity: 0.3;
        }
        
        .page-break {
            page-break-after: always;
            break-after: page;
            margin: 0;
            padding: 0;
            height: 0;
            clear: both;
            display: block;
        }
        
        /* Decorative elements */
        .corner-decoration {
            position: absolute;
            width: 60px;
            height: 60px;
            opacity: 0.1;
        }
        
        .corner-decoration.top-right {
            top: 10mm;
            right: 10mm;
            border-top: 3px solid var(--primary-blue);
            border-right: 3px solid var(--primary-blue);
        }
        
        .corner-decoration.bottom-left {
            bottom: 10mm;
            left: 10mm;
            border-bottom: 3px solid var(--primary-blue);
            border-left: 3px solid var(--primary-blue);
        }
    </style>
</head>
<body>

<!-- PAGE 1: Editorial Board -->
<div class="a4">
<div class="header">
    <div class="header-top">
        <div class="header-logo">
            {% if college_logo_path %}
            <img src="/{{ college_logo_path }}" alt="College Logo">
            {% elif college_logo_b64 %}
            <img src="data:image/png;base64,{{ college_logo_b64 }}" alt="College Logo">
            {% endif %}
        </div>
        <div class="header-center">
            {% if orbits_logo_path %}
            <img src="/{{ orbits_logo_path }}" alt="Orbits Logo">
            {% elif orbits_logo_b64 %}
            <img src="data:image/png;base64,{{ orbits_logo_b64 }}" alt="Orbits Logo">
            {% endif %}
            <p>DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING</p>
            <p>KGISL INSTITUTE OF TECHNOLOGY, COIMBATORE - 641035</p>
        </div>
        <div class="header-badge">
            {% if naac_badge_path %}
            <img src="/{{ naac_badge_path }}" alt="NAAC Badge">
            {% elif naac_badge_b64 %}
            <img src="data:image/png;base64,{{ naac_badge_b64 }}" alt="NAAC Badge">
            {% endif %}
        </div>
    </div>
    <div class="header-bottom">
        <div class="header-bottom-left">MONTHLY NEWSLETTER</div>
        <div class="header-bottom-center">{{ month }}-{{ year }}</div>
        <div class="header-bottom-right">{{ volume }}  {{ issue }}</div>
    </div>
</div>

{% if main_image %}
<div class="main-image">
    <img src="data:image/png;base64,{{ main_image }}" alt="Main Image">
</div>
{% endif %}

<div class="section-title">EDITORIAL BOARD</div>
<div class="section-underline">───────────────────────────────────────────────────</div>

<div class="editorial-board">
{% for member in editorial %}
    {% if member.Role %}
    <div class="board-member">
        <strong>{{ member.Role }}:</strong>
        {{ member.Name }} {{ member.Designation }}
    </div>
    {% endif %}
{% endfor %}
</div>
</div> <!-- End PAGE 1 -->

<!-- PAGE 2: Vision, Mission, PEO, PSO -->
<div class="page-break"></div>
<div class="a4">

<div class="header">
    <div class="header-top">
        <div class="header-logo">
            {% if college_logo_path %}
            <img src="/{{ college_logo_path }}" alt="College Logo">
            {% elif college_logo_b64 %}
            <img src="data:image/png;base64,{{ college_logo_b64 }}" alt="College Logo">
            {% endif %}
        </div>
        <div class="header-center">
            {% if orbits_logo_path %}
            <img src="/{{ orbits_logo_path }}" alt="Orbits Logo">
            {% elif orbits_logo_b64 %}
            <img src="data:image/png;base64,{{ orbits_logo_b64 }}" alt="Orbits Logo">
            {% endif %}
            <p>DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING</p>
            <p>KGISL INSTITUTE OF TECHNOLOGY, COIMBATORE - 641035</p>
        </div>
        <div class="header-badge">
            {% if naac_badge_path %}
            <img src="/{{ naac_badge_path }}" alt="NAAC Badge">
            {% elif naac_badge_b64 %}
            <img src="data:image/png;base64,{{ naac_badge_b64 }}" alt="NAAC Badge">
            {% endif %}
        </div>
    </div>
    <div class="header-bottom">
        <div class="header-bottom-left">MONTHLY NEWSLETTER</div>
        <div class="header-bottom-center">{{ month }}-{{ year }}</div>
        <div class="header-bottom-right">{{ volume }}  {{ issue }}</div>
    </div>
</div>

<!-- Vision / Mission / PEO / PSO combined two-column layout -->
<div style="display:flex; gap:18px; align-items:flex-start;">
    <div style="flex:1; min-width:0;">
        <div class="section-title red">VISION &amp; MISSION</div>
        <div style="display:flex; gap:20px; align-items:flex-start;">
            <div style="flex:1;">
                {% if vision %}
                    <h4 style="color:#003399; margin-bottom:8px; font-weight:700;">Vision</h4>
                    {% for item in vision %}
                        <div class="content" style="background:transparent; padding:0; margin-bottom:8px;">
                            <p style="color:#333;">{{ item.Content }}</p>
                        </div>
                    {% endfor %}
                {% endif %}

                {% if mission %}
                    <h4 style="color:#CC0000; margin-top:12px; margin-bottom:8px; font-weight:700;">Mission</h4>
                    <ul class="mission-list" style="margin-left:18px;">
                    {% for item in mission %}
                        <li style="margin-bottom:6px;">{{ item.Content }}</li>
                    {% endfor %}
                    </ul>
                {% endif %}
            </div>

            <div style="flex:0 0 40%; max-width:220px; text-align:center;">
                {% if vision_path %}
                    <img src="/{{ vision_path }}" alt="Vision" style="max-width:100%; height:auto; box-shadow:0 3px 8px rgba(0,0,0,0.12); border-radius:4px;"/>
                {% elif vision_b64 %}
                    <img src="data:image/png;base64,{{ vision_b64 }}" alt="Vision" style="max-width:100%; height:auto; box-shadow:0 3px 8px rgba(0,0,0,0.12); border-radius:4px;"/>
                {% endif %}
            </div>
        </div>

        <div style="margin-top:14px;">
            {% if peo %}
            <div class="page-break"></div>
            <div class="section-title">PROGRAM EDUCATIONAL OBJECTIVES (PEO'S)</div>
            <div class="section-underline">───────────────────────────────────────────────────</div>
            {% for item in peo %}
                {% if item.Code %}
                <div class="peo-item">
                    <strong>{{ item.Code }}:</strong>
                    <p>{{ item.Description }}</p>
                </div>
                {% endif %}
            {% endfor %}
            {% endif %}

            {% if pso %}
            <div class="section-title">PROGRAM SPECIFIC OUTCOMES (PSO'S)</div>
            <div class="section-underline">───────────────────────────────────────────────────</div>
            {% for item in pso %}
                {% if item.Code %}
                <div class="pso-item">
                    <strong>{{ item.Code }}:</strong>
                    <p>{{ item.Description }}</p>
                </div>
                {% endif %}
            {% endfor %}
            {% endif %}
        </div>
    </div>
</div>
</div> <!-- End PAGE 2 -->

<!-- PAGE 3: Table of Contents -->
<div class="page-break"></div>
<div class="a4">

<div class="header">
    <div class="header-top">
        <div class="header-logo">
            {% if college_logo_path %}
            <img src="/{{ college_logo_path }}" alt="College Logo">
            {% elif college_logo_b64 %}
            <img src="data:image/png;base64,{{ college_logo_b64 }}" alt="College Logo">
            {% endif %}
        </div>
        <div class="header-center">
            {% if orbits_logo_path %}
            <img src="/{{ orbits_logo_path }}" alt="Orbits Logo">
            {% elif orbits_logo_b64 %}
            <img src="data:image/png;base64,{{ orbits_logo_b64 }}" alt="Orbits Logo">
            {% endif %}
            <p>DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING</p>
            <p>KGISL INSTITUTE OF TECHNOLOGY, COIMBATORE - 641035</p>
        </div>
        <div class="header-badge">
            {% if naac_badge_path %}
            <img src="/{{ naac_badge_path }}" alt="NAAC Badge">
            {% elif naac_badge_b64 %}
            <img src="data:image/png;base64,{{ naac_badge_b64 }}" alt="NAAC Badge">
            {% endif %}
        </div>
    </div>
    <div class="header-bottom">
        <div class="header-bottom-left">MONTHLY NEWSLETTER</div>
        <div class="header-bottom-center">{{ month }}-{{ year }}</div>
        <div class="header-bottom-right">{{ volume }}  {{ issue }}</div>
    </div>
</div>

<div class="section-title">CONTENTS</div>
<div class="section-underline">───────────────────────────────────────────────────</div>

<table class="contents-table">
    <thead>
        <tr>
            <th style="width: 10%;">S.NO</th>
            <th style="width: 70%;">CONTENTS</th>
            <th style="width: 20%;">PAGE NO</th>
        </tr>
    </thead>
    <tbody>
        {% for idx, section_name in enumerate(sorted_sections) %}
            <tr>
                <td>{{ idx + 1 }}</td>
                <td style="text-align: center;">{{ section_name.upper() }}</td>
                <td>{{ section_page_map.get(section_name, 'XX') }}</td>
            </tr>
        {% endfor %}
    </tbody>
</table>
</div> <!-- End PAGE 3 -->

<!-- PAGES 4+: Event Sections -->
{% for section_name in sorted_sections %}
<div class="page-break"></div>
<div class="a4">

<div class="header">
    <div class="header-top">
        <div class="header-logo">
            {% if college_logo_path %}
            <img src="/{{ college_logo_path }}" alt="College Logo">
            {% elif college_logo_b64 %}
            <img src="data:image/png;base64,{{ college_logo_b64 }}" alt="College Logo">
            {% endif %}
        </div>
        <div class="header-center">
            {% if orbits_logo_path %}
            <img src="/{{ orbits_logo_path }}" alt="Orbits Logo">
            {% elif orbits_logo_b64 %}
            <img src="data:image/png;base64,{{ orbits_logo_b64 }}" alt="Orbits Logo">
            {% endif %}
            <p>DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING</p>
            <p>KGISL INSTITUTE OF TECHNOLOGY, COIMBATORE - 641035</p>
        </div>
        <div class="header-badge">
            {% if naac_badge_path %}
            <img src="/{{ naac_badge_path }}" alt="NAAC Badge">
            {% elif naac_badge_b64 %}
            <img src="data:image/png;base64,{{ naac_badge_b64 }}" alt="NAAC Badge">
            {% endif %}
        </div>
    </div>
    <div class="header-bottom">
        <div class="header-bottom-left">MONTHLY NEWSLETTER</div>
        <div class="header-bottom-center">{{ month }}-{{ year }}</div>
        <div class="header-bottom-right">{{ volume }}  {{ issue }}</div>
    </div>
</div>

<div class="section-title">{{ section_name.upper() }}</div>
<div class="section-underline">───────────────────────────────────────────────────</div>

{% for event in sections[section_name] %}
    <div class="event-title">{{ event['Event Title'] }}</div>
    
    {% if event['Event Date'] %}
    <div class="event-meta">Date: {{ event['Event Date'] }}</div>
    {% endif %}
    
    {% set event_details = event_details_map[loop.index0] %}
    {% if event_details %}
    <div class="event-meta">{{ event_details | join(' | ') }}</div>
    {% endif %}
    
    {% set image_ref = str(event['Image Reference']) if event['Image Reference'] else '' %}
    {% if image_ref in embedded_images %}
    <div class="event-image">
        <img src="data:image/png;base64,{{ embedded_images[image_ref] }}" alt="Event Image">
    </div>
    {% endif %}
    
    {% if event['Event Description'] %}
    <div class="event-description">
        {{ event['Event Description'] }}
    </div>
    {% endif %}
    
    {% if event['Coordinators'] %}
    <div class="event-meta">Coordinators: {{ event['Coordinators'] }}</div>
    {% endif %}
    
    <hr>
{% endfor %}
</div> <!-- End Event Section Page -->

{% endfor %}

<!-- FINAL PAGE: Contact -->
<div class="page-break"></div>
<div class="a4 last-page">

<div class="header">
    <div class="header-top">
        <div class="header-logo">
            {% if college_logo_path %}
            <img src="/{{ college_logo_path }}" alt="College Logo">
            {% elif college_logo_b64 %}
            <img src="data:image/png;base64,{{ college_logo_b64 }}" alt="College Logo">
            {% endif %}
        </div>
        <div class="header-center">
            {% if orbits_logo_path %}
            <img src="/{{ orbits_logo_path }}" alt="Orbits Logo">
            {% elif orbits_logo_b64 %}
            <img src="data:image/png;base64,{{ orbits_logo_b64 }}" alt="Orbits Logo">
            {% endif %}
            <p>DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING</p>
            <p>KGISL INSTITUTE OF TECHNOLOGY, COIMBATORE - 641035</p>
        </div>
        <div class="header-badge">
            {% if naac_badge_path %}
            <img src="/{{ naac_badge_path }}" alt="NAAC Badge">
            {% elif naac_badge_b64 %}
            <img src="data:image/png;base64,{{ naac_badge_b64 }}" alt="NAAC Badge">
            {% endif %}
        </div>
    </div>
    <div class="header-bottom">
        <div class="header-bottom-left">MONTHLY NEWSLETTER</div>
        <div class="header-bottom-center">{{ month }}-{{ year }}</div>
        <div class="header-bottom-right">{{ volume }}  {{ issue }}</div>
    </div>
</div>

<br><br>

<div class="section-title">CONTACT</div>
<div class="section-underline">───────────────────────────────────────────────────</div>

<div class="contact-info">
    <p><strong>Follow Us on:</strong></p>
    <div class="social-media">
        <span class="social-icon youtube" title="YouTube">
            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="M23.498 6.186a3.016 3.016 0 0 0-2.122-2.136C19.505 3.545 12 3.545 12 3.545s-7.505 0-9.377.505A3.017 3.017 0 0 0 .502 6.186C0 8.07 0 12 0 12s0 3.93.502 5.814a3.016 3.016 0 0 0 2.122 2.136c1.871.505 9.376.505 9.376.505s7.505 0 9.377-.505a3.015 3.015 0 0 0 2.122-2.136C24 15.93 24 12 24 12s0-3.93-.502-5.814zM9.545 15.568V8.432L15.818 12l-6.273 3.568z"/></svg>
        </span>
        <span class="social-icon instagram" title="Instagram">
            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="M12 2.163c3.204 0 3.584.012 4.85.07 3.252.148 4.771 1.691 4.919 4.919.058 1.265.069 1.645.069 4.849 0 3.205-.012 3.584-.069 4.849-.149 3.225-1.664 4.771-4.919 4.919-1.266.058-1.644.07-4.85.07-3.204 0-3.584-.012-4.849-.07-3.26-.149-4.771-1.699-4.919-4.92-.058-1.265-.07-1.644-.07-4.849 0-3.204.013-3.583.07-4.849.149-3.227 1.664-4.771 4.919-4.919 1.266-.057 1.645-.069 4.849-.069zm0-2.163c-3.259 0-3.667.014-4.947.072-4.358.2-6.78 2.618-6.98 6.98-.059 1.281-.073 1.689-.073 4.948 0 3.259.014 3.668.072 4.948.2 4.358 2.618 6.78 6.98 6.98 1.281.058 1.689.072 4.948.072 3.259 0 3.668-.014 4.948-.072 4.354-.2 6.782-2.618 6.979-6.98.059-1.28.073-1.689.073-4.948 0-3.259-.014-3.667-.072-4.947-.196-4.354-2.617-6.78-6.979-6.98-1.281-.059-1.69-.073-4.949-.073zm0 5.838c-3.403 0-6.162 2.759-6.162 6.162s2.759 6.163 6.162 6.163 6.162-2.759 6.162-6.163c0-3.403-2.759-6.162-6.162-6.162zm0 10.162c-2.209 0-4-1.79-4-4 0-2.209 1.791-4 4-4s4 1.791 4 4c0 2.21-1.791 4-4 4zm6.406-11.845c-.796 0-1.441.645-1.441 1.44s.645 1.44 1.441 1.44c.795 0 1.439-.645 1.439-1.44s-.644-1.44-1.439-1.44z"/></svg>
        </span>
        <span class="social-icon linkedin" title="LinkedIn">
            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433c-1.144 0-2.063-.926-2.063-2.065 0-1.138.92-2.063 2.063-2.063 1.14 0 2.064.925 2.064 2.063 0 1.139-.925 2.065-2.064 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z"/></svg>
        </span>
        <span class="social-icon twitter" title="X (Twitter)">
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M18.244 2.25h3.308l-7.227 8.26 8.502 11.24H16.17l-5.214-6.817L4.99 21.75H1.68l7.73-8.835L1.254 2.25H8.08l4.713 6.231zm-1.161 17.52h1.833L7.084 4.126H5.117z"/></svg>
        </span>
    </div>
    <p class="social-handle">@kitetechcollege</p>
    <p style="margin-top: 12px;"><strong>The Editor</strong></p>
    <p>Department of Computer Science And Engineering</p>
    <p>KGiSL Institute of Technology</p>
    <p>KGiSL Campus, 365, Thudiyalur Road,</p>
    <p>Saravanampatti, Coimbatore – 641035</p>
    <p>e-Mail: {{ contact.Email if contact.Email else 'kitecse@kgkite.ac.in' }}</p>
</div>

<div class="footer-board">
{% for member in editorial %}
    {% if member.Role and ('editor' in member.Role.lower() or 'managing' in member.Role.lower() or 'executive' in member.Role.lower() or 'director' in member.Role.lower()) %}
    <div class="footer-board-member">
        <div class="name">{{ member.Name }}</div>
        <div class="role">{{ member.Role }}</div>
    </div>
    {% endif %}
{% endfor %}
</div>

</div> <!-- .a4 -->

</body>
</html>
"""
        # Build event details for all events
        event_details_list = []
        for event in self.data['events']:
            event_details_list.append(self._build_event_details(event))
        
        sorted_sections = sorted(sections.keys())

        # Compute a simple page number map for the contents page.
        # Layout assumptions (simple mapping for preview/demo):
        # Page 1: Editorial Board
        # Page 2: Vision/Mission/PEO/PSO
        # Page 3: Contents
        # Page 4+ : one page per section (in the order of sorted_sections)
        start_section_page = 4
        section_page_map = {name: start_section_page + idx for idx, name in enumerate(sorted_sections)}
        
        template = Template(html_template)
        html_content = template.render(
            college_logo_b64=college_logo_b64,
            orbits_logo_b64=orbits_logo_b64,
            naac_badge_b64=naac_badge_b64,
            vision_b64=vision_b64,
            main_image=main_image,
            college_logo_path=college_logo_path,
            orbits_logo_path=orbits_logo_path,
            naac_badge_path=naac_badge_path,
            vision_path=vision_path,
            month=self.data['info'].get('Month', 'AUGUST'),
            year=self.data['info'].get('Year', '2024'),
            volume=self.data['info'].get('Volume', 'Volume 2'),
            issue=self.data['info'].get('Issue', 'Issue 1'),
            editorial=self.data['editorial'],
            vision=vision,
            mission=mission,
            peo=self.data['peo'],
            pso=self.data['pso'],
            sections=sections,
            sorted_sections=sorted_sections,
            section_page_map=section_page_map,
            events=self.data['events'],
            event_details_map=event_details_list,
            embedded_images=embedded_images,
            contact=self.data['contact'],
            enumerate=enumerate,
            str=str  # Pass Python's str function to template
        )
        
        return html_content
    
    def generate(self):
        """Generate HTML and PDF newsletter"""
        try:
            # Generate HTML
            html_content = self.generate_html()
            
            # Save HTML
            output_folder = os.path.join('generated', self.session_id)
            os.makedirs(output_folder, exist_ok=True)
            
            html_path = os.path.join(output_folder, 'newsletter.html')
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            return html_path
            
        except Exception as e:
            raise Exception(f"Error generating newsletter: {str(e)}")


def generate_html_newsletter(excel_path, image_paths, session_id):
    """Generate HTML-based newsletter"""
    generator = HTMLNewsletterGenerator(excel_path, image_paths, session_id)
    return generator.generate()
