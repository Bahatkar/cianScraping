package main

import (
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"strconv"
	"strings"
	"time"

	"github.com/spf13/viper"
	"github.com/xuri/excelize/v2"
)

type Flat struct {
	Id            int    `json:"id"`
	Rooms         int    `json:"roomsCount"`
	Square        string `json:"totalArea"`
	Floor         int    `json:"floorNumber"`
	Type          string `json:"offerType"`
	Apartment     bool   `json:"isApartment"`
	FromDeveloper bool   `json:"fromDeveloper"`
	House         struct {
		Section int `json:"section"`
	} `json:"house"`
	BargainTerms struct {
		Price int `json:"priceRur"`
	} `json:"bargainTerms"`
	Geo struct {
		JK struct {
			ComplexName string `json:"name"`
			URL         string `json:"fullUrl"`
			Building    struct {
				BuildingName string `json:"name"`
			} `json:"house"`
		} `json:"jk"`
	} `json:"geo"`
}

func main() {
	if err := initConfig(); err != nil {
		log.Fatalf("Error occured while reading config: %s", err)
	}

	var (
		url         = viper.GetString("url")
		rowCounter  = 4
		complexName string
	)

	excelFile := newExcelBook()
	for i := 1; ; i++ {
		//Собираю URL. Сортировка нужна для уменьшения шанса замены лота на дубль в выдаче
		sorting := "&sort=price_object_order"
		page := "&p=" + strconv.Itoa(i)
		reqURL := url + page + sorting
		resp, err := http.Get(reqURL)
		if err != nil {
			log.Fatalf("Error occured while requesting URL: %s\n", err)
		}

		/*
			Проверка на перенаправление: если произошло перенаправление на другую страницу, значит, запрашиваемого номера страницы не существует.
			Это означает конец сбора.
		*/
		if reqURL != resp.Request.URL.String() {
			log.Println("Парсинг завершен")
			break
		}

		log.Printf("Парсинг страницы %d\n", i)
		content, err := io.ReadAll(resp.Body)
		if err != nil {
			log.Println("Error occured while reading resp body")
		}
		//Обрезаю тело ответа с двух сторон, получая только JSON
		jsonPart1 := strings.Split(string(content), `"offers":[`)[1]
		jsonPart := "[" + strings.Split(jsonPart1, ",\"paginationUrls\"")[0]

		//Парсинг квартир из JSON и запись в эксель
		d := json.NewDecoder(strings.NewReader(jsonPart))
		d.Token()

		for d.More() {
			var flat Flat
			err = d.Decode(&flat)
			if err != nil {
				log.Println(err)
			}
			complexName, err = excelWriting(flat, rowCounter, excelFile)
			if err != nil {
				log.Printf("Error occured while writing excel: %s", err)
			}

			rowCounter++
		}
		d.Token()
		//Таймаут перед след. запросом для уменьшения шанса блокировки
		time.Sleep(5 * time.Second)
	}
	//Сохраняю файл эксель
	date := time.Now().Format("02.01.2006")
	fileName := fmt.Sprintf("%s_%s.%s", complexName, date, "xlsx")
	err := excelFile.SaveAs(fileName)
	if err != nil {
		log.Printf("Error occured while saving excel file: %s", err)
	}
}

// Создаю новую книгу эксель с шапкой по заданному шаблону
func newExcelBook() *excelize.File {
	var engCols = [17]string{"source_id", "url", "housing_complex_name", "building_name", "source", "rooms_number", "floor", "number_in_floor",
		"number_in_building", "section_number", "deal_type", "object_type", "square", "square_price", "overall_price", "date", "time"}
	var rusCols = [17]string{"п/п", "Название источника", "Название проекта", "Название корпуса", "Источник", "Количество комнат", "Этаж",
		"Номер на этаже", "Номер в корпусе", "Номер секции", "Тип сделки", "Тип объекта недвижимости", "Площадь", "Цена за метр(руб.)",
		"Общая цена (руб.)", "Дата", "Время"}

	book := excelize.NewFile()
	for i := 0; i < len(engCols); i++ {
		cellEng := fmt.Sprintf("%c1", 'A'+i)
		book.SetCellValue("Sheet1", cellEng, engCols[i])
		cellRus := fmt.Sprintf("%c2", 'A'+i)
		book.SetCellValue("Sheet1", cellRus, rusCols[i])
	}

	return book
}

// Запись собранной инфы в эксель
func excelWriting(ft Flat, rowCounter int, book *excelize.File) (string, error) {
	var (
		sh  = "Sheet1"
		url = strings.TrimSuffix(strings.TrimPrefix(ft.Geo.JK.URL, "https://"), "/")
	)
	if err := book.SetCellValue(sh, fmt.Sprintf("A%d", rowCounter), ft.Id); err != nil {
		return "", err
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("B%d", rowCounter), url); err != nil {
		return "", err
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("C%d", rowCounter), ft.Geo.JK.ComplexName); err != nil {
		return "", err
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("D%d", rowCounter), ft.Geo.JK.Building.BuildingName); err != nil {
		return "", err
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("E%d", rowCounter), "vp."+url); err != nil {
		return "", err
	}
	rooms := ft.Rooms
	if rooms > 4 {
		rooms = 4
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("F%d", rowCounter), rooms); err != nil {
		return "", err
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("G%d", rowCounter), ft.Floor); err != nil {
		return "", err
	}
	dealType := "Вторичка"
	if ft.FromDeveloper {
		dealType = "Первичка"
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("K%d", rowCounter), dealType); err != nil {
		return "", err
	}
	flatType := "Другое"
	if ft.Type == "flat" {
		flatType = "Квартира"
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("L%d", rowCounter), flatType); err != nil {
		return "", err
	}
	normSquare, _ := strconv.ParseFloat(ft.Square, 64)
	if err := book.SetCellValue(sh, fmt.Sprintf("M%d", rowCounter), normSquare); err != nil {
		return "", err
	}
	avg := getAveragePrice(ft.Square, ft.BargainTerms.Price)
	if err := book.SetCellValue(sh, fmt.Sprintf("N%d", rowCounter), avg); err != nil {
		return "", err
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("O%d", rowCounter), ft.BargainTerms.Price); err != nil {
		return "", err
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("P%d", rowCounter), time.Now().Format("02.01.2006")); err != nil {
		return "", err
	}
	if err := book.SetCellValue(sh, fmt.Sprintf("Q%d", rowCounter), time.Now().Format("15:04")); err != nil {
		return "", err
	}

	return ft.Geo.JK.ComplexName, nil
}

// Расчет средней цены за м2
func getAveragePrice(square string, price int) float64 {
	sq, _ := strconv.ParseFloat(square, 64)
	pr := float64(price)

	return pr / sq
}

// Инициализация конфигов из .yaml
func initConfig() error {
	viper.AddConfigPath("./")
	viper.SetConfigName("config")
	return viper.ReadInConfig()
}
