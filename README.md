# compareExcel
Private project - Simple script to compare 2 excel files

####
* Tested running node 12
* run `npm install` to install dependencies
* add `[file1].xlsx` and `[file1].xlsx` to root.
  * it uses the first sheet, this must contain the data.
  * important column A is uniue and files have headers.
* run `npm start name_of_file1 name_of_file2` -> this produces a `errorreport-[ID].xlsx` file.
  * if no file names are passed in it uses `file1` & `file2`