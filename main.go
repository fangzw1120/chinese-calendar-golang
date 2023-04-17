package main

import (
	"errors"
	"fmt"
	"git.woa.com/forisfang_tut/logger"
	"github.com/Lofanmi/chinese-calendar-golang/calendar"
	"github.com/alecthomas/kingpin"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

// User 不考虑闰月问题
var period = 10
var users = make([]*User, 0)

type User struct {
	ID             string
	Name           string
	Year           int64
	Month          int64
	Day            int64
	Birthday       string
	DateType       string
	FutureBirthday map[int64]string
}

const (
	UserIDHeader       = "工号"
	UserNameHeader     = "姓名"
	UserBirthHeader    = "生日"
	UserDateTypeHeader = "日历类型"
)

// GetMainDirectory 获取进程所在目录: 末尾带反斜杠
func GetMainDirectory() string {
	path, err := filepath.Abs(os.Args[0])

	if err != nil {
		return ""
	}

	fullPath := filepath.Dir(path)
	return pathAddBackslash(fullPath)
}

// PathAddBackslash 路径最后添加反斜杠
func pathAddBackslash(path string) string {
	i := len(path) - 1

	if !os.IsPathSeparator(path[i]) {
		path += string(os.PathSeparator)
	}
	return path
}

type Configure struct {
	Conf        *os.File `yaml:"-"`
	Addr        string   `yaml:"addr"`
	Port        int      `yaml:"port"`
	Root        string   `yaml:"root"`
	Prefix      string   `yaml:"prefix"`
	HTTPAuth    string   `yaml:"httpauth"`
	Cert        string   `yaml:"cert"`
	FileName    string
	OutFileName string
}

var gcfg = new(Configure)

func main() {
	kingpin.HelpFlag.Short('h')
	//kingpin.Flag("conf", "config file path, yaml format").FileVar(&gcfg.Conf)
	//kingpin.Flag("root", "root directory, default ./").Short('r').StringVar(&gcfg.Root)
	kingpin.Flag("FileName", "FileName").StringVar(&gcfg.FileName)
	kingpin.Flag("OutFileName", "OutFileName").StringVar(&gcfg.OutFileName)
	kingpin.Parse() // first parse conf

	logPath := GetMainDirectory()
	//"/Users/forisfang/project/github/chinese-calendar-golang/logs/"
	logger.Init(logPath, "tmp.log", true, false, true)
	//t := time.Now()
	// 1. ByTimestamp
	// 时间戳
	//c := calendar.ByTimestamp(t.Unix())
	// 2. BySolar
	// 公历
	//c := calendar.BySolar(year, month, day, hour, minute, second)
	// 3. ByLunar
	// 农历(最后一个参数表示是否闰月)
	//c := calendar.ByLunar(year, month, day, hour, minute, second, false)

	//filePath := "/Users/forisfang/Desktop/Book1.xlsx"
	filePath := GetMainDirectory() + gcfg.FileName
	header, sheetName, err := readExcelFile(filePath)
	if err != nil {
		logger.Errorf("%+v", err)
	}
	users = GetAllUser()

	nowYear := time.Now().Year()
	totYears := period
	yearLT := make([]int64, totYears)
	yearStrLt := make([]string, totYears)
	for i := 0; i < totYears; i++ {
		val := nowYear + i
		yearLT[i] = int64(val)
		yearStrLt[i] = strconv.Itoa(val) + "年"
	}
	logger.Debugf("calculate period: %+v", yearLT)

	// for user
	cnt := 0
	for _, user := range users {
		cnt++
		logger.Infof("user: %+v", user)
		futureBirthday := make(map[int64]string, totYears)
		for _, year := range yearLT {
			if user.DateType == "农历" {
				// 对应年份农历的时间对象
				valLunarItem := calendar.ByLunar(year, user.Month, user.Day, 0, 0, 0, false)
				// 对应年份农历
				//valLunar := valLunarItem.Lunar
				// 对应年份公历
				valSolar := valLunarItem.Solar
				futureBirthday[year] = fmt.Sprintf("%+v-%+v-%+v", valSolar.GetYear(), valSolar.GetMonthStr(), valSolar.GetDayStr())
				//logger.Debugf("年份: %+v, 农历: %+v-%+v-%+v, Solar: %+v-%+v-%+v", year,
				//	valLunar.GetYear(), valLunar.GetMonthStr(), valLunar.GetDayStr(),
				//	valSolar.GetYear(), valSolar.GetMonthStr(), valSolar.GetDayStr())

			} else if user.DateType == "阳历" {
				month, day := "", ""
				if user.Month <= 9 {
					month = "0" + strconv.Itoa(int(user.Month))
				} else {
					month = strconv.Itoa(int(user.Month))
				}
				if user.Day <= 9 {
					day = "0" + strconv.Itoa(int(user.Day))
				} else {
					day = strconv.Itoa(int(user.Day))
				}
				futureBirthday[year] = fmt.Sprintf("%+v-%+v-%+v", year, month, day)
			}
		}
		user.FutureBirthday = futureBirthday
		logger.Debugf("user: %+v", user)
	}
	logger.Infof("tot user: %+v", cnt)

	newHeader := append(header, yearStrLt...)
	writeToFile(GetMainDirectory()+gcfg.OutFileName, sheetName, newHeader, yearLT, users)

	//t := time.Now()
	//c := calendar.ByTimestamp(t.Unix())
	//lunarItem := c.Lunar
	//lunarStr := fmt.Sprintf("农历 %+v 年 %+v 月 %+v 日", lunarItem.GetYear(), lunarItem.GetMonth(), lunarItem.GetDay())
	//
	//bytes, err := c.ToJSON()
	//if err != nil {
	//	logger.Errorf("%+v", err)
	//}
	//logger.Debug(string(bytes))
	//logger.Debug(lunarStr)
}

func GetAllUser() []*User {
	return users
}

func readExcelFile(path string) ([]string, string, error) {
	header := make([]string, 0)
	sheetName := ""
	f, err := excelize.OpenFile(path)
	if err != nil {
		logger.Error(err.Error())
		return header, sheetName, err
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			logger.Error(err.Error())
		}
	}()

	// Get all the rows in the Sheet1.
	sheetNames := f.GetSheetList()
	if len(sheetNames) < 1 {
		logger.Error("sheet length error")
	}
	sheetName = sheetNames[0]
	logger.Debugf("fileName: %+v, sheetName: %+v", f.Path, sheetName)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		logger.Error(err.Error())
		return header, sheetName, err
	}

	idIdx, nameIdx, birthIdx, typeIdx := -1, -1, -1, -1
	for i, row := range rows {
		if i == 0 {
			header = row
			if idIdx, nameIdx, birthIdx, typeIdx = headerIndex(header); idIdx == -1 || nameIdx == -1 || birthIdx == -1 {
				msg := "header name error"
				logger.Errorf("%+v: %+v", msg, header)
				return header, sheetName, errors.New(msg)
			}
			logger.Infof("Header: %+v", header)
			continue
		}

		birthday := row[birthIdx]
		dateType := row[typeIdx]
		newBirthday := formatBirthday(f, row, sheetName, birthday, i)

		year, month, day := birthSplit(newBirthday)
		user := User{
			ID:       row[idIdx],
			Name:     row[nameIdx],
			Year:     year,
			Month:    month,
			Day:      day,
			Birthday: newBirthday,
			DateType: dateType,
		}
		logger.Debugf("user: %+v", user)
		users = append(users, &user)
	}
	return header, sheetName, nil
}

func birthSplit(birthday string) (int64, int64, int64) {
	sep := "-"
	year, month, day := int64(-1), int64(-1), int64(-1)
	ymd := strings.Split(birthday, sep)
	if len(ymd) >= 3 {
		i, err := strconv.Atoi(ymd[0])
		if err != nil || i < 1900 || i > 2100 {
			logger.Error("year value error")
		} else {
			year = int64(i)
		}
		i, err = strconv.Atoi(ymd[1])
		if err != nil || i <= 0 || i > 12 {
			logger.Error("year value error")
		} else {
			month = int64(i)
		}
		i, err = strconv.Atoi(ymd[2])
		if err != nil || i <= 0 || i > 31 {
			logger.Error("year value error")
		} else {
			day = int64(i)
		}
	}
	return year, month, day
}

func headerIndex(header []string) (int, int, int, int) {
	idIdx := indexOf(UserIDHeader, header)
	nameIdx := indexOf(UserNameHeader, header)
	birthIdx := indexOf(UserBirthHeader, header)
	dateTypeIdx := indexOf(UserDateTypeHeader, header)
	return idIdx, nameIdx, birthIdx, dateTypeIdx
}

func indexOf(element string, data []string) int {
	for k, v := range data {
		if element == v {
			return k
		}
	}
	return -1 //not found.
}

func indexOfInt(element int64, data []int64) int {
	for k, v := range data {
		if element == v {
			return k
		}
	}
	return -1 //not found.
}

func formatDate(f *excelize.File, sheetName string, cellName string) string {
	style, _ := f.NewStyle(&excelize.Style{NumFmt: 34, Lang: "ko-kr"})
	f.SetCellStyle(sheetName, cellName, cellName, style)
	e7, _ := f.GetCellValue(sheetName, cellName)
	return e7
}

func formatBirthday(f *excelize.File, row []string, sheetName, birthday string, i int) string {
	cellIdx := strconv.Itoa(i + 1)
	cellPre := toCharStr(indexOf(birthday, row) + 1)
	cellName := cellPre + cellIdx
	logger.Debugf("%+v, %+v, %+v, %+v th row, birthday cell: %+v", sheetName, row, birthday, i, cellName)
	return formatDate(f, sheetName, cellName)
}

func toCharStr(i int) string {
	return string('A' - 1 + i)
}

func checkResult() {

}

func writeToFile(filePath, sheetName string, newHeader []string, totYears []int64, users []*User) {
	f := excelize.NewFile() //creating a new sheet

	newSheetName := sheetName + "_result"
	idx, err := f.NewSheet(newSheetName) //creating the new sheet names
	if err != nil {
		logger.Errorf("%+v", err)
	}
	// set header
	for i, headerName := range newHeader {
		rowIdx := "1"
		prefix := toCharStr(i + 1)
		f.SetCellValue(newSheetName, prefix+rowIdx, headerName)
	}

	// set user
	for i, user := range users {
		rowIdx := strconv.Itoa(i + 2)
		f.SetCellValue(newSheetName, "A"+rowIdx, user.ID)
		f.SetCellValue(newSheetName, "B"+rowIdx, user.Name)
		f.SetCellValue(newSheetName, "C"+rowIdx, user.Birthday)
		f.SetCellValue(newSheetName, "D"+rowIdx, user.DateType)

		for j, year := range totYears {
			cellName := toCharStr(j+4) + rowIdx
			f.SetCellValue(newSheetName, cellName, user.FutureBirthday[year])
		}
	}

	f.SetActiveSheet(idx)
	if err := f.SaveAs(filePath); err != nil { //saving the new sheet in the file names companies
		logger.Errorf("%+v", err)
	}
}
