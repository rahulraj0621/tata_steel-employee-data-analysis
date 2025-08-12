import PyPDF2
import sys
import os
import json

def extract_pdf_text(pdf_path):
    if not os.path.exists(pdf_path):
        return f"Error: File not found - {pdf_path}"
    
    text_content = []
    
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            num_pages = len(reader.pages)
            
            for page_num in range(num_pages):
                page = reader.pages[page_num]
                text = page.extract_text()
                if text.strip():  
                    text_content.append({
                        "page": page_num + 1,
                        "content": text
                    })
            
            summary = {
                "filename": os.path.basename(pdf_path),
                "total_pages": num_pages,
                "pages_with_content": len(text_content)
            }
            
            return {
                "summary": summary,
                "content": text_content
            }
    except Exception as e:
        return f"Error extracting PDF: {str(e)}"

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = r"c:\Users\abhi8\Downloads\Coursera MASKED data.pdf"
    
    result = extract_pdf_text(pdf_path)
    
    if isinstance(result, dict):
        with open("pdf_extraction_result.json", "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
            
        with open("pdf_first_pages.txt", "w", encoding="utf-8") as f:
            for page in result["content"][:5]:
                f.write(f"\n\n==== PAGE {page['page']} ====\n\n")
                f.write(page["content"])
                
        print(f"Extraction complete. Found {result['summary']['total_pages']} pages, {result['summary']['pages_with_content']} with content.")
    else:
        print(result)  