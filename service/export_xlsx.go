package main

import (
	"database/sql"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	_ "github.com/lib/pq"
)

type postgreSQLConfig struct {
	Host     string
	Port     int
	User     string
	Password string
	Database string
}

type tableRow struct {
	id                int
	name              string
	email             string
	date              string
	phone             string
	phone2            string
	company           string
	siret             string
	rut_number        string
	personal_number   string
	org_number        string
	street_address    string
	city              string
	postal            string
	region            string
	country           string
	coordinates       string
	mastercard        string
	pin               string
	cvv               string
	track             string
	small_description string
	big_description   string
	alphanumeric      string
	currency          string
	status            string
	guid              string
	current_range     int
	accepted          string
	random_text       string
}

func checkError(err error) {
	if err != nil {
		panic(err)
	}
}

func connectPostgresDatabase() *sql.DB {
	configFile, err := ioutil.ReadFile("../database/database_config.json") // you need to create this file if it does not exist
	checkError(err)

	var config postgreSQLConfig
	err = json.Unmarshal(configFile, &config)
	checkError(err)

	psqlConnectionString := fmt.Sprintf(
		"host=%s port=%d user=%s password=%s dbname=%s sslmode=disable",
		config.Host, config.Port, config.User, config.Password, config.Database)

	db, err := sql.Open("postgres", psqlConnectionString)
	checkError(err)

	err = db.Ping()
	checkError(err)

	fmt.Println("Successfully connected to Postgres database")
	return db
}

func createXlsxWithHeader(sheetName string) *excelize.File {
	xlsxFile := excelize.NewFile()

	xlsxFile.MergeCell(sheetName, "A1", "C1")
	xlsxFile.SetColWidth(sheetName, "A", "AD", 30)

	style, err := xlsxFile.NewStyle(`{"font": {"bold": true}}`)
	checkError(err)
	xlsxFile.SetCellStyle(sheetName, "A1", "A1", style)
	xlsxFile.SetCellValue(sheetName, "A1", "Exported table example")

	style, err = xlsxFile.NewStyle(`{
		"alignment": {"horizontal": "center"},
		"font": {"bold": true},
		"border": [
		{
			"type": "left",
			"color": "000000",
			"style": 2
		},
		{
			"type": "top",
			"color": "000000",
			"style": 2
		},
		{
			"type": "bottom",
			"color": "000000",
			"style": 2
		},
		{
			"type": "right",
			"color": "000000",
			"style": 2
		}]
	}`)
	checkError(err)
	xlsxFile.SetCellStyle(sheetName, "A3", "AD3", style)
	xlsxFile.SetCellValue(sheetName, "A3", "id")
	xlsxFile.SetCellValue(sheetName, "B3", "name")
	xlsxFile.SetCellValue(sheetName, "C3", "email")
	xlsxFile.SetCellValue(sheetName, "D3", "date")
	xlsxFile.SetCellValue(sheetName, "E3", "phone")
	xlsxFile.SetCellValue(sheetName, "F3", "phone2")
	xlsxFile.SetCellValue(sheetName, "G3", "company")
	xlsxFile.SetCellValue(sheetName, "H3", "siret")
	xlsxFile.SetCellValue(sheetName, "I3", "rut_number")
	xlsxFile.SetCellValue(sheetName, "J3", "personal_number")
	xlsxFile.SetCellValue(sheetName, "K3", "org_number")
	xlsxFile.SetCellValue(sheetName, "L3", "street_address")
	xlsxFile.SetCellValue(sheetName, "M3", "city")
	xlsxFile.SetCellValue(sheetName, "N3", "postal")
	xlsxFile.SetCellValue(sheetName, "O3", "region")
	xlsxFile.SetCellValue(sheetName, "P3", "country")
	xlsxFile.SetCellValue(sheetName, "Q3", "coordinates")
	xlsxFile.SetCellValue(sheetName, "R3", "mastercard")
	xlsxFile.SetCellValue(sheetName, "S3", "pin")
	xlsxFile.SetCellValue(sheetName, "T3", "cvv")
	xlsxFile.SetCellValue(sheetName, "U3", "track")
	xlsxFile.SetCellValue(sheetName, "V3", "small_description")
	xlsxFile.SetCellValue(sheetName, "W3", "big_description")
	xlsxFile.SetCellValue(sheetName, "X3", "alphanumeric")
	xlsxFile.SetCellValue(sheetName, "Y3", "currency")
	xlsxFile.SetCellValue(sheetName, "Z3", "status")
	xlsxFile.SetCellValue(sheetName, "AA3", "guid")
	xlsxFile.SetCellValue(sheetName, "AB3", "current_range")
	xlsxFile.SetCellValue(sheetName, "AC3", "accepted")
	xlsxFile.SetCellValue(sheetName, "AD3", "random_text")

	return xlsxFile
}

func exportXlsx(xlsxFile *excelize.File, sheetName string, rowsToExport int) {
	db := connectPostgresDatabase()
	defer db.Close()

	fmt.Println("Starting export data from database")

	sql := fmt.Sprintf("SELECT * FROM import_table LIMIT %d;", rowsToExport)
	rows, err := db.Query(sql)
	checkError(err)

	rowIdx := 4
	for rows.Next() {
		row := tableRow{}
		err := rows.Scan(
			&row.id,
			&row.name,
			&row.email,
			&row.date,
			&row.phone,
			&row.phone2,
			&row.company,
			&row.siret,
			&row.rut_number,
			&row.personal_number,
			&row.org_number,
			&row.street_address,
			&row.city,
			&row.postal,
			&row.region,
			&row.country,
			&row.coordinates,
			&row.mastercard,
			&row.pin,
			&row.cvv,
			&row.track,
			&row.small_description,
			&row.big_description,
			&row.alphanumeric,
			&row.currency,
			&row.status,
			&row.guid,
			&row.current_range,
			&row.accepted,
			&row.random_text,
		)
		checkError(err)

		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIdx), row.id)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIdx), row.name)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("C%d", rowIdx), row.email)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("D%d", rowIdx), row.date)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("E%d", rowIdx), row.phone)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("F%d", rowIdx), row.phone2)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("G%d", rowIdx), row.company)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("H%d", rowIdx), row.siret)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("I%d", rowIdx), row.rut_number)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("J%d", rowIdx), row.personal_number)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("K%d", rowIdx), row.org_number)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("L%d", rowIdx), row.street_address)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("M%d", rowIdx), row.city)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("N%d", rowIdx), row.postal)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("O%d", rowIdx), row.region)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("P%d", rowIdx), row.country)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("Q%d", rowIdx), row.coordinates)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("R%d", rowIdx), row.mastercard)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("S%d", rowIdx), row.pin)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("T%d", rowIdx), row.cvv)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("U%d", rowIdx), row.track)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("V%d", rowIdx), row.small_description)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("W%d", rowIdx), row.big_description)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("X%d", rowIdx), row.alphanumeric)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("Y%d", rowIdx), row.currency)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("Z%d", rowIdx), row.status)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("AA%d", rowIdx), row.guid)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("AB%d", rowIdx), row.current_range)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("AC%d", rowIdx), row.accepted)
		xlsxFile.SetCellValue(sheetName, fmt.Sprintf("AD%d", rowIdx), row.random_text)
		rowIdx++
	}

	err = xlsxFile.SaveAs("../export/export.xlsx")
	checkError(err)
}

func main() {
	startTime := time.Now()
	fmt.Println(startTime.Format(time.UnixDate))

	sheetName := "Sheet1"
	xlsxFile := createXlsxWithHeader(sheetName)
	rowsToExport := 300000
	exportXlsx(xlsxFile, sheetName, rowsToExport)

	duration := time.Since(startTime)
	fmt.Printf("Export %d rows finished. Duration: %s \n", rowsToExport, duration)
}
