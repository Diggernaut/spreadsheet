package spreadsheet

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"net/url"

	"github.com/spf13/cast"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
)

const (
	baseURL = "https://sheets.googleapis.com/v4"

	// Scope is the API scope for viewing and managing your Google Spreadsheet data.
	// Useful for generating JWT values.
	Scope = "https://spreadsheets.google.com/feeds"

	// SecretFileName is used to get client.
	SecretFileName = "client_secret.json"
)

// NewService makes a new service with the secret file.
func NewService() (s *Service, err error) {
	data, err := ioutil.ReadFile(SecretFileName)
	if err != nil {
		return
	}

	conf, err := google.JWTConfigFromJSON(data, Scope)
	if err != nil {
		return
	}

	s = NewServiceWithClient(conf.Client(oauth2.NoContext))
	return
}

// NewServiceWithClient makes a new service by the client.
func NewServiceWithClient(client *http.Client) *Service {
	return &Service{
		baseURL: baseURL,
		client:  client,
	}
}

// Service represents a Sheets API service instance.
// Service is the main entry point into using this package.
type Service struct {
	baseURL string
	client  *http.Client
}

// CreateSpreadsheet creates a spreadsheet with the given title
func (s *Service) CreateSpreadsheet(spreadsheet Spreadsheet) (resp Spreadsheet, err error) {
	body, err := s.post("/spreadsheets", map[string]interface{}{
		"properties": map[string]interface{}{
			"title": spreadsheet.Properties.Title,
		},
	})
	if err != nil {
		return
	}
	err = json.Unmarshal([]byte(body), &resp)
	if err != nil {
		return
	}
	return s.FetchSpreadsheet(resp.ID)
}

// FetchSpreadsheet fetches the spreadsheet by the id.
func (s *Service) FetchSpreadsheet(id string) (spreadsheet Spreadsheet, err error) {
	fields := "spreadsheetId,properties.title,sheets(properties,data.rowData.values(formattedValue))"
	fields = url.QueryEscape(fields)
	path := fmt.Sprintf("/spreadsheets/%s?fields=%s", id, fields)
	body, err := s.get(path)
	if err != nil {
		return
	}
	err = json.Unmarshal(body, &spreadsheet)
	if err != nil {
		return
	}
	spreadsheet.service = s
	return
}
func (s *Service) NewSheet(sh *Sheet, l int, name string) (err error) {
	path := "https://sheets.googleapis.com/v4/spreadsheets/" + sh.Spreadsheet.ID + ":batchUpdate"
	d := []byte(`
		[{
		  "addSheet": {
			"properties": {
			  "title": "` + name + `",
			  "gridProperties": {
				"rowCount": 0,
				"columnCount": 0
			  }
			}
		}
	  }]`)
	var p = make([]map[string]interface{}, 0)
	err = json.Unmarshal(d, &p)
	if err != nil {
		return err
	}
	var e = make(map[string]interface{})
	e["requests"] = p
	body, err := sh.Spreadsheet.service.postv4(path, e)
	if err != nil {
		return err
	}
	fmt.Println(body)
	return
}
func (s *Service) ClearSheet(sh *Sheet) (err error) {
	path := "https://sheets.googleapis.com/v4/spreadsheets/" + sh.Spreadsheet.ID + ":batchUpdate"
	// path := "https://sheets.googleapis.com/v4/spreadsheets/" + sh.Spreadsheet.ID + "/values:clear"

	d := []byte(`[
		{
		  "updateCells": {
			"range": {
			  "sheetId": ` + cast.ToString(sh.Properties.ID) + `
			},
			"fields": "*"
		  }
		}
	  ]`)
	var p = make([]map[string]interface{}, 0)
	json.Unmarshal(d, &p)
	var e = make(map[string]interface{})
	e["requests"] = p
	body, err := sh.Spreadsheet.service.postv4(path, e)
	fmt.Println(body)
	return
}

// SyncSheet updates sheet
func (s *Service) SyncSheet(sheet *Sheet) (err error) {
	if sheet.newMaxRow > sheet.Properties.GridProperties.RowCount ||
		sheet.newMaxColumn > sheet.Properties.GridProperties.ColumnCount {
		err = s.ExpandSheet(sheet, sheet.newMaxRow, sheet.newMaxColumn)
		if err != nil {
			return
		}
	}
	err = s.syncCells(sheet)
	if err != nil {
		return
	}
	sheet.modifiedCells = []*Cell{}
	sheet.newMaxRow = sheet.Properties.GridProperties.RowCount
	sheet.newMaxColumn = sheet.Properties.GridProperties.ColumnCount
	return
}
func (s *Service) SyncRawSheet(sheet *Sheet, maxrow, maxcol uint, cell []*Cell) (err error) {
	log.Println("here1")
	// if maxrow > sheet.newMaxRow ||
	// 	maxcol > sheet.newMaxColumn {
	// 	log.Println("here2")
	err = s.ExpandSheet(sheet, maxrow, maxcol)
	if err != nil {
		return
	}
	// }
	log.Println("here2")
	err = s.syncRawCells(sheet, cell)
	if err != nil {
		return
	}
	log.Println("here3")
	sheet.modifiedCells = []*Cell{}
	log.Println("here4")
	sheet.newMaxRow = sheet.Properties.GridProperties.RowCount
	log.Println("here5")
	sheet.newMaxColumn = sheet.Properties.GridProperties.ColumnCount
	log.Println("here6")
	return
}
func (s *Service) syncRawCells(sheet *Sheet, cell []*Cell) (err error) {
	path := fmt.Sprintf("/spreadsheets/%s/values:batchUpdate", sheet.Spreadsheet.ID)
	params := map[string]interface{}{
		"valueInputOption": "USER_ENTERED",
		"data":             make([]map[string]interface{}, 0, len(cell)),
	}
	for _, cell := range cell {
		valueRange := map[string]interface{}{
			"range":          sheet.Properties.Title + "!" + cell.Pos(),
			"majorDimension": "COLUMNS",
			"values": [][]string{
				[]string{
					cell.Value,
				},
			},
		}
		params["data"] = append(params["data"].([]map[string]interface{}), valueRange)
	}
	log.Println("hah1")
	// _, err = sheet.Spreadsheet.service.post(path, params)
	_, err = s.post(path, params)

	log.Println("hah2")
	return
}

// ExpandSheet expands the range of the sheet
func (s *Service) ExpandSheet(sheet *Sheet, row, column uint) (err error) {
	props := sheet.Properties
	props.GridProperties.RowCount = row
	props.GridProperties.ColumnCount = column
	log.Println("hh")
	log.Println(row)
	log.Println(column)
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	log.Println("hh2")
	err = r.UpdateSheetProperties(sheet, &props, false).DoService(s)
	if err != nil {
		return
	}
	log.Println("hh3")
	sheet.newMaxRow = row
	sheet.newMaxColumn = column
	return
}

// DeleteRows deletes rows from the sheet
func (s *Service) DeleteRows(sheet *Sheet, start, end int) (err error) {
	sheet.Properties.GridProperties.RowCount -= uint(end - start)
	sheet.newMaxRow -= uint(end - start)
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.DeleteDimension(sheet, "ROWS", start, end).Do()
	return
}

// DeleteColumns deletes columns from the sheet
func (s *Service) DeleteColumns(sheet *Sheet, start, end int) (err error) {
	sheet.Properties.GridProperties.ColumnCount -= uint(end - start)
	sheet.newMaxRow -= uint(end - start)
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.DeleteDimension(sheet, "COLUMNS", start, end).Do()
	return
}

func (s *Service) syncCells(sheet *Sheet) (err error) {
	path := fmt.Sprintf("/spreadsheets/%s/values:batchUpdate", sheet.Spreadsheet.ID)
	params := map[string]interface{}{
		"valueInputOption": "USER_ENTERED",
		"data":             make([]map[string]interface{}, 0, len(sheet.modifiedCells)),
	}
	for _, cell := range sheet.modifiedCells {
		valueRange := map[string]interface{}{
			"range":          sheet.Properties.Title + "!" + cell.Pos(),
			"majorDimension": "COLUMNS",
			"values": [][]string{
				[]string{
					cell.Value,
				},
			},
		}
		params["data"] = append(params["data"].([]map[string]interface{}), valueRange)
	}
	_, err = sheet.Spreadsheet.service.post(path, params)
	return
}

func (s *Service) get(path string) (body []byte, err error) {
	resp, err := s.client.Get(baseURL + path)
	if err != nil {
		return
	}
	body, err = ioutil.ReadAll(resp.Body)
	resp.Body.Close()
	if err != nil {
		return
	}
	err = s.checkError(body)
	return
}
func (s *Service) postv4(path string, params interface{}) (body string, err error) {
	reqBody, err := json.Marshal(params)
	if err != nil {
		return
	}
	resp, err := s.client.Post(path, "application/json", bytes.NewReader(reqBody))
	if err != nil {
		return
	}
	bytes, err := ioutil.ReadAll(resp.Body)
	resp.Body.Close()
	if err != nil {
		return
	}
	err = s.checkError(bytes)
	if err != nil {
		return
	}
	body = string(bytes)
	return
}
func (s *Service) post(path string, params interface{}) (body string, err error) {
	reqBody, err := json.Marshal(params)
	if err != nil {
		return
	}
	resp, err := s.client.Post(baseURL+path, "application/json", bytes.NewReader(reqBody))
	if err != nil {
		return
	}
	bytes, err := ioutil.ReadAll(resp.Body)
	resp.Body.Close()
	if err != nil {
		return
	}
	err = s.checkError(bytes)
	if err != nil {
		return
	}
	body = string(bytes)
	return
}

func (s *Service) checkError(body []byte) (err error) {
	var res map[string]interface{}
	err = json.Unmarshal(body, &res)
	if err != nil {
		return
	}
	resErr, hasErr := res["error"].(map[string]interface{})
	if !hasErr {
		return
	}
	code := resErr["code"].(float64)
	message := resErr["message"].(string)
	status := resErr["status"].(string)
	if err != nil {
		err = fmt.Errorf("error status: %s, code:%d, message: %s", status, int(code), message)
		return
	}
	return nil
}
