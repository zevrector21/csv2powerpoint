# CSV2PPTX

Load data, images, audio and video into Powerpoint slides using python

## Requirements
* VS Code (useful for debugging) or run from CLI
* `python` 3.x
* Powerpoint
* LibreOffice
* ffmpeg

## Setup
* Create venv in this folder
* Activate and install requirements.txt with pip
* Open folder in VS Code
* Select python interpreter.venv/Scripts/python.exe
* Run demo.py

## Helpful Links
* <https://python-pptx.readthedocs.io/en/latest/>
* <https://code.visualstudio.com/docs/python/python-tutorial>
* <https://www.libreoffice.org/download/download/>
* <https://www.ffmpeg.org/>

## How To Create Slide Templates
<https://docs.google.com/document/d/1lK88kGAeDpTva42BTo3_ua1LmjaGhS4jR6iTJ1w-hlc/edit?usp=sharing>


## Common Fixes when setting up a machine
*  Save your own screenshot for the save button
    - You need to do this if LibreOffice isn't finding a save button
    - After doing this you may still have to check 'always save' the first time
    - You may need to go to the progam and back to your terminal when running

Mac/Linux
* Linux, the program was not able to read the path  os.path.exists('inputs\save.png')  [Missing file save.png error (resolved)]
    - This needs changed to be:  `os.path.exists('inputs/save.png')` in every file. 

* Linux:  update this line in all files :   `args = ['C:\Program Files\LibreOffice\program\soffice', '--impress', '--nologo', ppt_output_filename]`
`args = ['loimpress', '--nologo', ppt_output_filename]. (Make sure the libreoffice is installed and it can be opened in the terminal by just writing:  loimpress )`
