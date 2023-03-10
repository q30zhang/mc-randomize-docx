# mc-randomize-docx

Randomize the order of choices (not questions) of multiple choice questions, in docx format.

## Format of the input file
- a *.docx file of multiple choice questions.
- one empty paragraph (empty line) should present between questions.
- an empty paragraph should present in the end of document.
- the correct anwser (support only single correct answer) of each question should use a different color.
- labels for the questions and the choices do not matter (1. 2. 3. or i. ii. iii.; A) B) or a. b. etc.).
- an example file is provided.

## Output files
- outputs are docx files with re-ordered choices and corresponding answer keys in txt files.
- output docx files still have the correct answers colored differently; double check the answers and reset all color to black before use.
- output files are in the same directory as the input file.

## Requirements
- python-docx >= 0.8.11

## Usage
```
python randomize_choice.py "<file_path>" <number of versions> <number of choices in each question>
```
