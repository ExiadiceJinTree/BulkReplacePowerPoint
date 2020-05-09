- [日本語版(In Japanese)](./README_JP.md)

# Overview
Replaces the target character string in the Power Point file that contains the specified replacement source string regular expression in the specified folder and subordinate folders with the specified replacement destination string.

# Description
- As a Power Point file, both .ppt and .pptx are supported.
- Search(pre-check) and replace
  - Search(pre-check):  
    Before performing the replacement, this can output the folder path and the file name of the Power Point file containing the replacement source string, and the text containing the replacement source string in the file as a TSV file, therefoe you can check files and texts which will be replaced in advance.
    - The output TSV file
      - Output destination folder path: Folder where the execution script exists.
      - File name pattern: SearchResult_YYYYMMDDhhmmss.tsv
  - You can choose whether to execute the search(pre-check) process and the replace process as you like.
- About specifying the search/replacement target character string
  - Please note that it is treated as a regular expression.
  - It is not case sensitive.

# Usage
Run the BulkReplacePowerPoint.vbs file and follow the prompt dialogs that appear.
