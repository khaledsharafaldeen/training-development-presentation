To generate the .pptx file locally (instructions)

1) Clone the repository:
   git clone https://github.com/khaledsharafaldeen/training-development-presentation.git

2) Enter the folder and create a virtual environment:
   cd training-development-presentation/presentation
   python -m venv venv
   source venv/bin/activate    # Windows: venv\Scripts\activate

3) Install dependency:
   pip install python-pptx

4) Run the generator script:
   python generate_pptx.py

5) The generated file will be in presentation/output/Training_Development_Strategy_Khaled_Sharaf_Aldeen.pptx

Notes:
- The script uses the font name "Noto Sans Arabic". For best results, install this font on your system before opening the PPTX.
- After generating, you can edit the file in PowerPoint to adjust layouts, add logo, or tweak styling.
