package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/tealeg/xlsx"
)

func main() {
	files, err := filepath.Glob("*")
	if err != nil {
		log.Fatal(err)
	}

	fname := getContains(files, ".xlsx")

	if fname == "" {
		return
	}

	xlFile, err := xlsx.OpenFile(fname)
	if err != nil {
		fmt.Println("Não foi possível abrir o arquivo informado.")
		log.Fatalln(err.Error())
		return
	}

	parseSheets(xlFile.Sheets)
}

func parseSheets(sheets []*xlsx.Sheet) {
	for _, sheet := range sheets {
		for i, row := range sheet.Rows {
			if i == 0 {
				continue
			}

			fname := row.Cells[7].Value
			fname = createFile(fname, i)

			f := openFile(fname)
			defer func() {
				if err := f.Sync(); err != nil {
					fmt.Println("Não foi possível modificar o arquivo.")
				}
				f.Close()
			}()

			for celli, cell := range row.Cells {
				text := cell.String()
				key := sheet.Rows[0].Cells[celli].String()

				if key == "ID" || key == "Email" || key == "Nome" || key == "Hora de início" || key == "Hora de conclusão" {
					continue
				}

				writeToFile(f, key, text)
			}
		}
	}
}

func getContains(strs []string, pattern string) string {
	for _, s := range strs {
		if strings.Contains(s, pattern) {
			return s
		}
	}

	return ""
}

func createFile(path string, i int) string {
	if strings.TrimSpace(path) == "" {
		path = fmt.Sprintf("sem_nome_%d", i)
	}
	path = path + ".txt"

	var _, err = os.Stat(path)

	if os.IsNotExist(err) {
		var file, err = os.Create(path)
		if err != nil {
			fmt.Println("Os arquivos existentes serão substituídos.")
			return path
		}

		defer file.Close()
	}

	return path
}

func openFile(path string) *os.File {
	var file, err = os.OpenFile(path, os.O_RDWR, 0644)
	if err != nil {
		fmt.Println("Não é possível escrever no arquivo")
		return nil
	}

	return file
}

func writeToFile(f *os.File, key, value string) {
	_, _ = f.WriteString(fmt.Sprintf("%s: %s\n", key, value))
}
