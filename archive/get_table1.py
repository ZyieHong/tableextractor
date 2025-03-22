from pptx import Presentation
import pandas as pd

def extract_slide_data(ppt_path):
    # Extract slide data (title n tables)
    prs = Presentation(ppt_path)
    slide_info = []
    
    for slide_num, slide in enumerate(prs.slides, start=1):
        # Extract title
        title = extract_slide_title(slide)
        
        # Extract and process tables
        tables = process_slide_tables(slide)
        
        slide_info.append({
            "slide_number": slide_num,
            "title": title,
            "tables": tables,
            "table_count": len(tables)
        })
        
        print(f"Slide {slide_num}: {title} [{len(tables)} Table]")
    
    return slide_info

def extract_slide_title(slide):
    # Extract slide title (optimized version)
    if slide.shapes.title and slide.shapes.title.text.strip():
        return slide.shapes.title.text.strip()
    
    text_shapes = sorted(
        [s for s in slide.shapes if s.has_text_frame],
        key=lambda x: (x.top, x.left)
    )
    
    for shape in text_shapes:
        if shape.text.strip():
            return shape.text.strip()
    
    return "No Title"

def process_slide_tables(slide):
    # Process all tables in a single slide 
    tables = []
    for shape in slide.shapes:
        if shape.has_table:
            table_data = process_table(shape.table)
            df = table_to_dataframe(table_data)
            tables.append(df)
    return tables

def process_table(table):
    # Process the complete data of a single table
    rows = len(table.rows)
    cols = len(table.columns)
    grid = [['' for _ in range(cols)] for _ in range(rows)]
    
    # Handle merged cells
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            if cell.is_merge_origin:
                value = cell.text.strip()
                row_span = cell.span_height
                col_span = cell.span_width
                
                # Fill the merged area
                for x in range(i, i + row_span):
                    for y in range(j, j + col_span):
                        if x < rows and y < cols:
                            grid[x][y] = value
            else:
                # Regular cells take values directly
                grid[i][j] = cell.text.strip()
    
    return grid

def table_to_dataframe(table_data):
    # Convert to dF
    try:
        if not table_data:
            return pd.DataFrame()
        
        # Automatically identify headers (first row)
        headers = table_data[0] if table_data else []
        
        # Process data rows (exclude possible empty rows)
        data_rows = [row for row in table_data[1:] if any(cell != '' for cell in row)]
        
        # Create DataFrame
        df = pd.DataFrame(data_rows, columns=headers)
        
        # Clean up leading/trailing spaces
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
        
        return df
    except Exception as e:
        print(f"Table conversion error: {str(e)}")
        return pd.DataFrame(table_data)

if __name__ == "__main__":
    
    keyword = input("Keyword: ").strip()
    slides = extract_slide_data(r"C:\Users\User\OneDrive - ump.edu.my\BSD UMP\DSP1111\tests\test_data\cimb.pptx")
    # User interaction
    results = [slide for slide in slides if keyword.lower() in slide["title"].lower()]
    
    if not results:
        print("No Table Found")
    else:
        print(f"\nSlide: {[s['slide_number'] for s in results]}")
        
        for res in results:
            print(f"\n{'='*80}")
            print(f"Slide #{res['slide_number']}")
            print(f"Title: {res['title']}")
            
            if res['table_count'] > 0:
                for idx, df in enumerate(res['tables'], 1):
                    print(f"\nTable {idx} ({df.shape[1]} columns x {df.shape[0]} rows)")
                    print(df.to_string(index=False))
                    print("-"*50)
            else:
                print("\nNo Table Found!")