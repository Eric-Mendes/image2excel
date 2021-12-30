import os

import numpy as np
import pandas as pd
from openpyxl import load_workbook, styles
from PIL import Image

# Loading the image as a PIL Image
image_name = ""  # it should have the extension with it
img = Image.open(f"images/{image_name}").convert("RGB")

# Resizing it and getting it as a 2d list with the RGB colors of each pixel
img = img.resize((img.size[0] // 10, img.size[1] // 5))
arr = np.array(img).tolist()
arr = [["%02x%02x%02x" % tuple(item) for item in row] for row in arr]

df = pd.DataFrame(arr)
# the same that was defined above but without extension
image_name = os.path.splitext(image_name)[0]

# Saving a DataFrame where each cell has a text corresponding to the RGB color its background should be
df.to_excel(f"spreadsheets/{image_name}.xlsx", index=False, header=False)


# Loading the excel file, painting each cell with its color and saving the updates
wb = load_workbook(f"spreadsheets/{image_name}.xlsx")

ws = wb.active
ws.title = image_name

for row in range(1, df.shape[0] + 1):
    for col in range(1, df.shape[1] + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = styles.PatternFill(
            start_color=cell.value, 
            end_color=cell.value, 
            fill_type="solid"
        )

wb.save(f"spreadsheets/{image_name}.xlsx")