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
	"github.com/tealeg/xlsx"
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

var rowChan = make(chan []string, 1024)

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
	rowChan <- columns
	for rows.Next() {
		err = rows.Scan(scans...)
		if err != nil {
			fmt.Println("Failed to scan row", err)
			// return results
		}
		row := make([]string, len(columns))
		for i := range vals {
			row[i] = string(vals[i])
		}
		// results = append(results, row)
		// fmt.Println("row:", row)
		rowChan <- row
	}
	close(rowChan)
	/*
		if len(results) == 1 {
			warning := []string{"no data found!"}
			results = append(results, warning)
		}
	*/
	// return results
}

func ExportData(stop chan bool) {
	fmt.Println("Export starts")

	streamf := new(xlsx.StreamFile)
	// var file *xlsx.File
	var tab *xlsx.Sheet
	var row *xlsx.Row
	var err error
	//if the file already exists
	var defaultFontSize = 11
	var defaultFontName = "Calibri"
	xlsx.SetDefaultFont(defaultFontSize, defaultFontName)
	if utils.Exists(fp) {
		streamf.xlsxFile, err = xlsx.OpenFile(fp)
		if err != nil {
			panic(err.Error())
		}
		//if sheet exists
		if _, ok := streamf.xlsxFile.Sheet[sheet]; ok {
			// tab = file.Sheet[sheet]
			panic("This sheet already exists!")
		} else {
			tab, err = streamf.xlsxFile.AddSheet(sheet)
			if err != nil {
				fmt.Printf(err.Error())
			}
		}

	} else {
		streamf.xlsxFile = xlsx.NewFile()
		tab, err = streamf.xlsxFile.AddSheet(sheet)
		if err != nil {
			fmt.Printf(err.Error())
		}
	}
	/*
		for i := range data {
			row = tab.AddRow()
			row.WriteSlice(&data[i], -1)
		}
	*/
	for r := range rowChan {
		row = tab.AddRow()
		// fmt.Println("r:", r)
		row.WriteSlice(&r, -1)
	}

	err = streamf.xlsxFile.Save(fp)
	if err != nil {
		fmt.Printf(err.Error())
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
