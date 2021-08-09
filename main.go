package main

import (
	"ExportToExcel/utils"
	"database/sql"
	"flag"
	"fmt"
	"log"
	"os"
	"time"

	_ "github.com/denisenkom/go-mssqldb"
	excelize "github.com/xuri/excelize/v2"
)

/*
// var isdebug = true
// var server = "PTPC-39PWGQ2\\SQL2016"
// var port = 1433
// var user = "sa"
// var password = "Salary.com"
// var database = "DB1"
ExportToExcel.exe -s "PTPC-39PWGQ2\SQL2016" -port 1433 -d db1 -u sa -p Salary.com -t tbl_Post -fp D:\Study\GO\src\github.com\ken0912\studygo\ExportToExcel\tbl_Post.xlsx -sheet Sheet5

ExportToExcel.exe -s PTPC-39PWGQ2\SQL2016 -d DB1 -t tbl_Post -fp tbl_Post_20200508.xlsx -sheet tbl_Post

*/

var (
	h       bool
	v, V    bool
	isdebug bool
	s, S    string
	port    int
	u, U    string
	p, P    string
	d, D    string
	t       string
	q, Q    string
	fp      string
	sheet   string
)

func init() {
	flag.BoolVar(&h, "h", false, "this help")
	flag.BoolVar(&v, "v", false, "show version and exit")
	flag.BoolVar(&v, "V", false, "show version and configure options then exit")
	flag.BoolVar(&isdebug, "isdebug", false, "show db connection info")
	flag.StringVar(&s, "s", "", "db server name ")
	flag.StringVar(&s, "S", "", "db server name ")
	flag.IntVar(&port, "port", 1433, "db port")
	flag.StringVar(&u, "u", "", "db login user name")
	flag.StringVar(&u, "U", "", "db login user name")
	flag.StringVar(&p, "p", "", "db login password")
	flag.StringVar(&p, "P", "", "db login password")
	flag.StringVar(&d, "d", "", "database name")
	flag.StringVar(&d, "D", "", "database name")
	flag.StringVar(&t, "t", "", "table name")
	flag.StringVar(&q, "q", "", "Query string")
	flag.StringVar(&q, "Q", "", "Query string")
	flag.StringVar(&fp, "fp", "", "full path of the export data")
	flag.StringVar(&sheet, "sheet", "Sheet1", "sheet name of the excel file")

}
func usage() {
	fmt.Fprintf(os.Stderr, `ExportDataToExcel Tool version: windows/0.0.1
Usage: ExportToExcel [-hvVsSuUpPdDtfpsheet] [-s signal] [-c filename] [-p prefix] [-g directives]

Options:
`)
	flag.PrintDefaults()
}

var rowChan = make(chan []interface{}, 1024)

func GetResult() {
	//连接字符串
	var connString string
	var sqlstr string
	/*
		if u == "" || U == "" {
			connString = fmt.Sprintf("server=%s;port%d;trusted_connection=yes;database=%s", s, port, d)
		} else {
			connString = fmt.Sprintf("server=%s;port%d;database=%s;user id=%s;password=%s", s, port, d, u, p)
		}
	*/
	connString = fmt.Sprintf("server=%s;port%d;database=%s;user id=%s;password=%s", s, port, d, u, p)
	if isdebug {
		fmt.Println(connString)
	}
	//建立连接
	db, err := sql.Open("mssql", connString)
	if err != nil {
		log.Fatal("Open Connection failed:", err.Error())
	}
	defer db.Close()

	//通过连接对象执行查询
	if t == "" {
		sqlstr = q
	} else {
		sqlstr = "select * from " + t
	}

	// fmt.Println("sqlstr:", sqlstr)
	rows, err := db.Query(sqlstr)

	if err != nil {
		log.Fatal("Query failed:", err.Error())
	}
	defer rows.Close()
	columns, err := rows.Columns()
	columnsInterface := make([]interface{}, len(columns))
	for i, column := range columns {
		columnsInterface[i] = column
	}
	if err != nil {
		log.Fatalln(err)
	}

	vals := make([][]byte, len(columns))
	scans := make([]interface{}, len(columns))

	for i := range vals {
		scans[i] = &vals[i]

	}

	// var results [][]string
	// results = append(results, columns)
	rowChan <- columnsInterface
	for rows.Next() {
		err = rows.Scan(scans...)
		if err != nil {
			fmt.Println("Failed to scan row", err)
			// return results
		}
		row := make([]interface{}, len(columns))
		for i := range vals {
			row[i] = string(vals[i])
		}
		// results = append(results, row)
		// fmt.Println("row:", row)
		rowChan <- row
	}
	close(rowChan)

}

func ExportData(stop chan bool) {
	fmt.Println("Export starts")

	var file *excelize.File
	var streamWriter *excelize.StreamWriter
	var err error
	if utils.Exists(fp) {
		file, err = excelize.OpenFile(fp)
		if err != nil {
			panic(err.Error())
		}
		if sheetIndex := file.GetSheetIndex(sheet); sheetIndex == -1 {
			if idx := file.NewSheet(sheet); idx == -1 {
				fmt.Println("NewSheet error")
				return
			}
			streamWriter, err = file.NewStreamWriter(sheet)
			if err != nil {
				fmt.Println("NewStreamWriter err:", err)
				return
			}
		} else {
			panic("This sheet already exists!")
		}

	} else {
		file = excelize.NewFile()

		file.DeleteSheet(sheet)

		if idx := file.NewSheet(sheet); idx == -1 {
			fmt.Println("idx:", idx)
			fmt.Println("NewSheet error")
			return
		}
		if sheet == "sheet1" {
			sheet = "Sheet1"
		}
		streamWriter, err = file.NewStreamWriter(sheet)
		if err != nil {
			fmt.Println(err)
		}
	}
	rowID := 1
	for row := range rowChan {
		cell, _ := excelize.CoordinatesToCellName(1, rowID)
		if err := streamWriter.SetRow(cell, row); err != nil {
			fmt.Println(err)
		}
		rowID += 1
	}
	if err := streamWriter.Flush(); err != nil {
		fmt.Println(err)
	}
	if err := file.SaveAs(fp); err != nil {
		fmt.Println(err)
	}
	stop <- true
}
func validation() {
	if t == "" && q == "" {
		panic("-t and -q can not both allowed to be empty")
	}
	if fp == "" {
		fmt.Println("-fp:", fp)
		panic("-fp is not allowed to be empty")
	}
	if len(sheet) >= 31 {
		sheet = sheet[0:31]
		// panic("sheet name is to long")
	}

}

func main() {

	flag.Parse()

	if h {
		flag.Usage()
		return
	}
	validation()

	stop := make(chan bool)
	go GetResult()
	go ExportData(stop)

	for {
		select {
		case <-stop:
			fmt.Println("Export Done!")
			return
		default:
			time.Sleep(1 * time.Second)
		}
	}
}
