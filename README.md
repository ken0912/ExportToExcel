# ExportToExcel
a commond line tool for Export Data from ssms to excel
use :ExportToExcel -h in cmd model to show all params of it
you can use it for example:
ExportToExcel.exe -s "PTPC-39PWGQ2\SQL2016" -port 1433 -d db1 -u sa -p sa12345 -t tbl_Post -fp D:\Study\GO\tbl_Post.xlsx -sheet Sheet5

ExportToExcel.exe -s PTPC-39PWGQ2\SQL2016 -d DB1 -t tbl_Post -fp tbl_Post_20200508.xlsx -sheet tbl_Post

tips:
it supports multiple tab exports
