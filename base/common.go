package base

import "strconv"

type Common struct {
	Month, Day int64
}

// GetMonthStr 月
func (common *Common) GetMonthStr() string {
	if common.Month <= 9 {
		return "0" + strconv.Itoa(int(common.Month))
	}
	return strconv.Itoa(int(common.Month))
}

// GetDayStr 日
func (common *Common) GetDayStr() string {
	if common.Day <= 9 {
		return "0" + strconv.Itoa(int(common.Day))
	}
	return strconv.Itoa(int(common.Day))
}
