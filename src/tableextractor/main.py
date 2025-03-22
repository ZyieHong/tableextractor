from pptx import Presentation
import pandas as pd

def extract_slide_data(ppt_path):
    prs = Presentation(ppt_path)
    slide_info = []
    
    for slide_num, slide in enumerate(prs.slides, start=1):
        title = _extract_slide_title(slide)
        tables = _process_slide_tables(slide)
        
        slide_info.append({
            "slide_number": slide_num,
            "title": title,
            "tables": tables,
            "table_count": len(tables)
        })
    
    return slide_info

def _extract_slide_title(slide):
    """Extract slide title (internal method)"""
    # Prefer using title placeholder
    if slide.shapes.title and slide.shapes.title.text.strip():
        return slide.shapes.title.text.strip()
    
    # Fallback: Sort text shapes by position
    text_shapes = sorted(
        [s for s in slide.shapes if s.has_text_frame],
        key=lambda x: (x.top, x.left)
    )
    
    for shape in text_shapes:
        if shape.text.strip():
            return shape.text.strip()
    
    return "No Title"

def _process_slide_tables(slide):
    """Process all tables in a single slide (internal method)"""
    tables = []
    for shape in slide.shapes:
        if shape.has_table:
            table_data = _process_table(shape.table)
            df = _table_to_dataframe(table_data)
            tables.append(df)
    return tables

def _process_table(table):
    """Process the full data of a single table (internal method)"""
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
                # Regular cells take values directly (avoid overwriting merged cells)
                if not grid[i][j]:
                    grid[i][j] = cell.text.strip()
    
    return grid

def _table_to_dataframe(table_data):
    """Convert to standard DataFrame (internal method)"""
    try:
        if not table_data or len(table_data) < 1:
            return pd.DataFrame()
        
        # Automatically identify headers (first non-empty row)
        header_row = next(
            (i for i, row in enumerate(table_data) if any(cell.strip() for cell in row)),
            None
        )
        
        if header_row is None:
            return pd.DataFrame()
        
        headers = [cell.strip() for cell in table_data[header_row]]
        
        # Process data rows (exclude possible empty rows)
        data_rows = [
            [cell.strip() for cell in row]
            for row in table_data[header_row+1:]
            if any(cell.strip() for cell in row)
        ]
        
        df = pd.DataFrame(data_rows, columns=headers)
        
        # Clean up blank values
        df.replace('', pd.NA, inplace=True)
        df.dropna(how='all', inplace=True)
        
        return df
    
    except Exception as e:
        print(f"Table conversion error: {str(e)}")
        return pd.DataFrame()