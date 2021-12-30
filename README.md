# image2excel
Draws an image on Excel

![Screenshot from 2021-11-12 13-15-02](https://user-images.githubusercontent.com/42689328/141498894-4ba72e43-930c-420c-b171-92d7dbfe6d91.png)

## Setting up the development environment
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
1. Get any image you want and put it inside the `images/` folder;
2. Correctly pass the path to the image_name string on this line: 
```python
image_name = ""  # it should have the extension with it
```
3. Lower the dimensions a little so that nothing breaks (<strong>a big image might break the code or your excel when you try to open the output file</strong>):
```python
img = img.resize((img.size[0]//6, img.size[1]//2))
```
*What's the best dimensions for your image? You have to find that out by trying different numbers when you resize it. We still don't have a nice way of checking that :(*

4. Run the code (keep in mind that you need to be inside the pipenv shell):
```bash
python3 main.py
```
The output will be inside the `spreadsheets/` folder and it'll be a `.xlsx` file with the same name of your image.

Have fun!
