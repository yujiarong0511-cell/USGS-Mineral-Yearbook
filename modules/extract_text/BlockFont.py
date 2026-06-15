# Input: A block with format info (var name: block_info, data type: nested dictionary) from .get_text("dict") in PyMuPDF
# Output: for block_format() function, the output is a dictionary about block's layout properties

import numpy as np
## --------------------Integrated Solution-------------------
def block_format(block_info):
    ## Block Index
    index_in_page = block_info['number'] # The index acquired by purely apply page.get_text("dict") or page.get_text("blocks") without sorting
    upperleft_x_coordinate = block_info['bbox'][0]
    upperleft_y_coordinate = block_info['bbox'][1]
    lowerright_x_coordinate = block_info['bbox'][2]
    lowerright_y_coordinate = block_info['bbox'][3]

    ## Font Size & Style 
    # Initialize
    font_size_list = [] # Each item is the font size of a span
    font_style_list = [] # Each item is the font style of a span
    spans_length_list = [] # Each item is the text length of the span

    is_firstline = True # Indicator of whether at the first line 
    is_firstline_indented = None
    firstline_font_size_list = []
    firstline_font_style_list = []
    firstline_spans_length_list = []

    # Iterate
    for line in block_info['lines']: # Iterate lines in the block
        if is_firstline:
            firstline_indentation = line['bbox'][0]
            is_firstline_indented = firstline_indentation > upperleft_x_coordinate # If the first line start at the right of the upper-left corner of the block, then the block is first-line-indented.
        for span in line['spans']: # Iterate spans in the line
            current_font_size = span['size'] # Extract font size of the span 
            current_font_style = span['font'] # Extract font style of the span 
            current_span_length = len(span['text']) # Calculate text length of the span 
            font_size_list.append(current_font_size)
            font_style_list.append(current_font_style)
            spans_length_list.append(current_span_length)
            if is_firstline:
                firstline_font_size_list.append(current_font_size)
                firstline_font_style_list.append(current_font_style)
                firstline_spans_length_list.append(current_span_length)
        is_firstline = False
    
    # Extract font size with highest frequency
    freq_table_size = zip(spans_length_list, font_size_list) # IMPORTANT: Order of the two parameters CANNOT be inverted
    dominant_font_size = max(freq_table_size)[1] # Extract the second element (font size) of the ordered 2-tuple (first element is the frequency)
    dominant_font_size = round(dominant_font_size) # Round to the closest integer (to avoid calculation errors when comparing integers and floats)

    # Extract font style with highest frequency
    freq_table_style = zip(spans_length_list, font_style_list) # IMPORTANT: Order of the two parameters CANNOT be inverted
    dominant_font_style = max(freq_table_style)[1] # Extract the second element (font size) of the ordered 2-tuple (first element is the frequency)

    ## Identify categories where the block belongs to
    # Whether is header: dominant_font_size > 10
    is_header = dominant_font_size > 10
    # Whether is footer: dominant_font_size < 10 + layout feature (extends very wide horizontally and locates at the bottom part of the page)
    is_footer = dominant_font_size < 10 and upperleft_x_coordinate < 50 and lowerright_x_coordinate > 550 and upperleft_y_coordinate > 700
    # Whether is footnote: (1) First span of the block is a numeric superscript; (2) dominant_font_size < 10 (font size of main text)
    is_footnote = block_info['lines'][0]['spans'][0]['size'] < 6 and dominant_font_size < 10
    # Whether is reference:
    is_reference = dominant_font_size < 10 and not (is_footer or is_footnote)
    # Whether is main-body (excluding reference): 
    is_mainbody = dominant_font_size == 10
    # Whether is section-title in main-text (a subset of main-body blocks, not possible WITHIN reference section):
    is_sec_title = dominant_font_size == 10 and "Bold" in dominant_font_style and "Ital" not in dominant_font_style
    # Whether is section-subtitle in main-text (a subset of main-body blocks, not possible WITHIN reference section):
    is_sec_subtitle = dominant_font_size == 10 and "Bold" in dominant_font_style and "Ital" in dominant_font_style
    
    output = {
        "index_in_page": index_in_page,
        "is_firstline_indented": is_firstline_indented,
        "dominant_font_size": dominant_font_size,
        "dominant_font_style": dominant_font_style,
        "is_header": is_header,
        "is_footer": is_footer,
        "is_footnote": is_footnote,
        "is_reference": is_reference,
        "is_mainbody": is_mainbody,
        "is_sec_title": is_sec_title,
        "is_sec_subtitle": is_sec_subtitle,
        "upperleft_x_coordinate": upperleft_x_coordinate,
        "upperleft_y_coordinate": upperleft_y_coordinate,
        "lowerright_x_coordinate": lowerright_x_coordinate,
        "lowerright_y_coordinate": lowerright_y_coordinate
    }
    return output


## ---------------------Font-Size-related---------------------
def AvgFontSize(block_info):
    # Initialize
    font_size_list = [] # Each item is the font size of a span
    text_length_list = [] # Each item is the text length of the span

    # Iterate
    for line in block_info['lines']: # Iterate lines in the block
        for span in line['spans']: # Iterate spans in the line
            current_font_size = span['size'] # Extract font size of the span  
            current_text_length = len(span['text']) # Calculate text length of the span 
            font_size_list.append(current_font_size)
            text_length_list.append(current_text_length)
    
    # Prepare for weight vector
    total_text_length = sum(text_length_list)
    font_text_weight = [text_length/total_text_length for text_length in text_length_list]
    # Calculate weighed average and round to closest integer
    average_font_size = round(float(np.dot(font_size_list, font_text_weight)))

    return average_font_size