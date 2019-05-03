# compareExcel
Private project - Simple script to compare 2 excel files

####
* Tested running node 12
* run `npm install` to install dependencies
* add `[file1].xlsx` and `[file1].xlsx` to root.
  * it uses the first sheet with data
  * important column A is unique and files have headers.
     * There is option for using column A an B as ID, se below
* run one of these
  * `npm run single name_of_file1 name_of_file2`
    * uses only column A as ID
    * this produces a `errorReport-[ID].xlsx` file.
  * `npm run double name_of_file1 name_of_file2`
    * uses only column A and B as ID
    * this produces a `errorReport-[ID].xlsx` file.

If no file names are passed in it uses `file1` & `file2`