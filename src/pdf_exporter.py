"""
PDF export functionality for 90D-PPT Generator
Converts PowerPoint presentations to PDF format
"""

import tempfile
import os
from pathlib import Path
from typing import Optional
import logging

# Try to import PDF conversion libraries
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.units import inch
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

try:
    import comtypes.client
    COMTYPES_AVAILABLE = True
except ImportError:
    COMTYPES_AVAILABLE = False


class PDFExporter:
    """Export PowerPoint presentations to PDF format"""
    
    def __init__(self):
        self.logger = logging.getLogger("PDFExporter")
    
    def is_pdf_export_available(self) -> bool:
        """Check if PDF export functionality is available"""
        return REPORTLAB_AVAILABLE or COMTYPES_AVAILABLE
    
    def get_pdf_export_method(self) -> str:
        """Get the available PDF export method"""
        if COMTYPES_AVAILABLE and os.name == 'nt':
            return "powerpoint"  # Use PowerPoint COM automation on Windows
        elif REPORTLAB_AVAILABLE:
            return "reportlab"  # Use ReportLab to recreate slides
        else:
            return "none"
    
    def export_to_pdf_via_powerpoint(self, ppt_path: str, output_path: str) -> bool:
        """Export PowerPoint to PDF using PowerPoint COM automation (Windows only)"""
        if not COMTYPES_AVAILABLE or os.name != 'nt':
            return False
        
        try:
            # Initialize PowerPoint application
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1
            
            # Open the presentation
            presentation = powerpoint.Presentations.Open(ppt_path)
            
            # Export to PDF
            # ppFixedFormatTypePDF = 2
            presentation.ExportAsFixedFormat(output_path, 2)
            
            # Close and cleanup
            presentation.Close()
            powerpoint.Quit()
            
            self.logger.info(f"Successfully exported PDF via PowerPoint: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to export via PowerPoint: {str(e)}")
            return False
    
    def export_to_pdf_via_reportlab(self, data_by_ranking: dict, rows_per_slide: int, output_path: str) -> bool:
        """Create PDF directly using ReportLab (cross-platform)"""
        if not REPORTLAB_AVAILABLE:
            return False
        
        try:
            doc = SimpleDocTemplate(output_path, pagesize=A4)
            styles = getSampleStyleSheet()
            story = []
            
            # Title page
            title = Paragraph("90-Day Planning PowerPoint", styles['Title'])
            story.append(title)
            story.append(Spacer(1, 20))
            
            # Process each ranking in order
            ranking_order = ['Accepted', 'Up Next', 'Maybe', 'Likely No']
            slide_num = 1
            
            for ranking in ranking_order:
                if ranking not in data_by_ranking or data_by_ranking[ranking].empty:
                    continue
                
                df = data_by_ranking[ranking]
                
                # Calculate number of slides for this ranking
                import math
                num_slides = math.ceil(len(df) / rows_per_slide)
                
                for slide_idx in range(num_slides):
                    start_row = slide_idx * rows_per_slide
                    end_row = min(start_row + rows_per_slide, len(df))
                    slide_data = df.iloc[start_row:end_row]
                    
                    # Add page break if not first slide
                    if slide_num > 1:
                        story.append(Spacer(1, 20))
                    
                    # Slide title
                    slide_title = f"{ranking} JIRA"
                    if num_slides > 1:
                        slide_title += f" (Page {slide_idx + 1} of {num_slides})"
                    
                    story.append(Paragraph(slide_title, styles['Heading1']))
                    story.append(Paragraph("Will be worked this 90", styles['Normal']))
                    story.append(Spacer(1, 12))
                    
                    # Status summary
                    closed_count = 0
                    non_closed_count = 0
                    if '_status_category' in slide_data.columns:
                        closed_count = (slide_data['_status_category'] == 'closed').sum()
                        non_closed_count = (slide_data['_status_category'] == 'non-closed').sum()
                    
                    status_text = f"In Progress/On Hold/Pending Review: {non_closed_count}, Closed: {closed_count}"
                    story.append(Paragraph(status_text, styles['Normal']))
                    story.append(Spacer(1, 12))
                    
                    # Create data table
                    table_data = [['JIRA', 'Target Complete Date', 'Description', 'Status', 'Risks/Issues', 'Component']]
                    
                    for _, row in slide_data.iterrows():
                        table_row = [
                            str(row.get('jira', '')),
                            '',  # Target date is empty
                            str(row.get('description', ''))[:50] + ('...' if len(str(row.get('description', ''))) > 50 else ''),
                            str(row.get('status', '')),
                            str(row.get('risks', ''))[:30] + ('...' if len(str(row.get('risks', ''))) > 30 else ''),
                            str(row.get('component', ''))
                        ]
                        table_data.append(table_row)
                    
                    # Create table
                    table = Table(table_data, colWidths=[1*inch, 1.2*inch, 2.5*inch, 0.8*inch, 1.5*inch, 1*inch])
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.black),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (-1, -1), 8),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ]))
                    
                    story.append(table)
                    story.append(Spacer(1, 12))
                    
                    # Page number
                    story.append(Paragraph(f"P{slide_num}", styles['Normal']))
                    
                    slide_num += 1
            
            # Build PDF
            doc.build(story)
            
            self.logger.info(f"Successfully created PDF via ReportLab: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to create PDF via ReportLab: {str(e)}")
            return False
    
    def export_presentation_to_pdf(self, ppt_path: str = None, data_by_ranking: dict = None, 
                                 rows_per_slide: int = 10) -> Optional[str]:
        """
        Export presentation to PDF using the best available method
        
        Args:
            ppt_path: Path to PowerPoint file (for PowerPoint method)
            data_by_ranking: Data dictionary (for ReportLab method)
            rows_per_slide: Number of rows per slide
            
        Returns:
            Path to created PDF file, or None if export failed
        """
        if not self.is_pdf_export_available():
            self.logger.error("No PDF export methods available")
            return None
        
        # Create temporary PDF file
        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_pdf.close()
        pdf_path = temp_pdf.name
        
        export_method = self.get_pdf_export_method()
        success = False
        
        if export_method == "powerpoint" and ppt_path:
            success = self.export_to_pdf_via_powerpoint(ppt_path, pdf_path)
        elif export_method == "reportlab" and data_by_ranking:
            success = self.export_to_pdf_via_reportlab(data_by_ranking, rows_per_slide, pdf_path)
        
        if success and Path(pdf_path).exists():
            return pdf_path
        else:
            # Cleanup failed attempt
            try:
                os.unlink(pdf_path)
            except:
                pass
            return None
    
    def get_export_info(self) -> dict:
        """Get information about available PDF export methods"""
        return {
            'available': self.is_pdf_export_available(),
            'method': self.get_pdf_export_method(),
            'reportlab_available': REPORTLAB_AVAILABLE,
            'comtypes_available': COMTYPES_AVAILABLE,
            'platform': os.name
        }