# image2excel
Draws an image on Excel

![image](https://user-images.githubusercontent.com/42689328/148095762-8e1b83e4-8923-49ab-bfdd-6d1c75af6450.png)

## Setting up the environment
Assuming you have `python >= 3.9` and `pipenv` installed:
```bash
# Cloning the repo 
git clone https://github.com/Eric-Mendes/image2excel.git
cd image2excel/

# Opening the environment and installing the dependencies
pipenv shell
pipenv install
```

## Running the project
Essentially all you need is to get your image ready and tweak the parameters on the `image2excel/config.ini` file. No manipulation of the `main.py` file is necessary.

1. Get any image you want and put it inside the `images/` folder <strong>OR</strong> copy the url of an image on the internet;

    - <strong>IMPORTANT:</strong> passing an url will download the image into the `images/` folder. If you think that some url looks sketchy, do <strong>not</strong> use it.

2. Correctly pass the path or the url to the image string: 

    - `image=this_is_fine.png` <strong>or</strong>
    - `image=https://media.geeksforgeeks.org/wp-content/uploads/20210318103632/gfg-300x300.png` (*just an example of an url*)

3. Lower the dimensions a little by a factor so that nothing breaks (<strong>a big image might break the code or your excel when you try to open the output file</strong>):

    - `factor=5`

        - *What's the best factor for your image? You have to find that out by trying different numbers when you resize it. We still don't have a nice way of checking that :(*

4. `row_height` and `column_width` were specifically set to make the cells a square so it looks even more like a pixel, but those hardcoded numbers might not look good on your screen since different computers have different viewports. You can tweak them if necessary.

5. `cell_value` assumes True or False. By default it is False, which means: do not display any text on the cells.

6. `zoom_scale` is the spreadsheet's zoom scale. By default it is 20 - the max zoom out possible - but you can also tweak this parameter if necessary. Keep in mind that it has to be an integer.

7. From source, run the code (keep in mind that you need to be inside the pipenv shell):
```bash
python3 image2excel/main.py
```
The output will be inside the `spreadsheets/` folder and it'll be a `.xlsx` file with the same name of your image.

Have fun!

- *Inspired by Matt Parker's [Stand-up comedy routine about Spreadsheets](https://youtu.be/UBX2QQHlQ_I).*
