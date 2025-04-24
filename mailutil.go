package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"os/exec"
	"sort"
	"strconv"
	"strings"
	"time"
)

type Recipient struct {
	Name     string
	Position string
	Email    string
	Birthday string
}

type BirthdayPerson struct {
	Name     string
	Position string
	Month    int
	Day      int
}

func GetBirthdayPeople(today string, recipients []Recipient) []BirthdayPerson {
	var result []BirthdayPerson
	for _, r := range recipients {
		if r.Birthday == "" {
			continue
		}
		parts := strings.Split(r.Birthday, "-")
		if len(parts) < 3 {
			continue
		}
		birthdayKey := fmt.Sprintf("%02s-%02s", parts[1], parts[2])
		if birthdayKey == today {
			result = append(result, BirthdayPerson{
				Name:     r.Name,
				Position: r.Position,
			})
		}
	}
	return result
}

func GetNextBirthdays(list []Recipient, count int) []BirthdayPerson {
	type birthdayWithDate struct {
		person BirthdayPerson
		diff   int
	}

	now := time.Now()
	var upcoming []birthdayWithDate

	for _, r := range list {
		if r.Birthday == "" {
			continue
		}
		parts := strings.Split(r.Birthday, "-")
		if len(parts) < 3 {
			continue
		}
		month, _ := strconv.Atoi(parts[1])
		day, _ := strconv.Atoi(parts[2])

		bdThisYear := time.Date(now.Year(), time.Month(month), day, 0, 0, 0, 0, time.Local)
		if bdThisYear.Before(now) {
			bdThisYear = bdThisYear.AddDate(1, 0, 0)
		}
		diff := int(bdThisYear.Sub(now).Hours() / 24)

		upcoming = append(upcoming, birthdayWithDate{
			person: BirthdayPerson{
				Name:     r.Name,
				Position: r.Position,
				Month:    month,
				Day:      day,
			},
			diff: diff,
		})
	}

	sort.Slice(upcoming, func(i, j int) bool {
		return upcoming[i].diff < upcoming[j].diff
	})

	var result []BirthdayPerson
	for i := 0; i < len(upcoming) && i < count; i++ {
		result = append(result, upcoming[i].person)
	}
	return result
}

func LoadRecipients(filename string) ([]Recipient, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	reader := csv.NewReader(file)
	headers, err := reader.Read()
	if err != nil {
		return nil, err
	}

	// 정확한 한국어 열 이름만 찾음
	colMap := map[string]int{
		"성명":     -1,
		"직위":     -1,
		"E-mail": -1,
		"생년월일":   -1,
	}

	for i, h := range headers {
		header := strings.TrimSpace(h)
		if _, ok := colMap[header]; ok {
			colMap[header] = i
		}
	}

	for k, v := range colMap {
		if v == -1 {
			return nil, fmt.Errorf("'%s' 열을 찾을 수 없습니다", k)
		}
	}

	var list []Recipient
	for {
		record, err := reader.Read()
		if err != nil {
			break
		}

		r := Recipient{
			Name:     strings.TrimSpace(record[colMap["성명"]]),
			Position: strings.TrimSpace(record[colMap["직위"]]),
			Email:    strings.TrimSpace(record[colMap["E-mail"]]),
			Birthday: strings.TrimSpace(record[colMap["생년월일"]]),
		}

		if r.Email != "" {
			list = append(list, r)
		}
	}

	if len(list) == 0 {
		return nil, fmt.Errorf("수신자 목록이 비어 있습니다")
	}

	return list, nil
}

func PrintRecipients(list []Recipient) {
	fmt.Println("수신자 목록:")
	for _, r := range list {
		fmt.Printf(" - %s (%s), 생일: %s, 이메일: %s\n", r.Name, r.Position, r.Birthday, r.Email)
	}
}

func FmtTime(t time.Time) string {
	return fmt.Sprintf("%02d%02d", t.Hour(), t.Minute())
}

func IsDigits(s string) bool {
	for _, r := range s {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}

func SendMail(toList []Recipient, birthdayPeople []BirthdayPerson) {
	if len(birthdayPeople) == 0 {
		fmt.Println("생일자가 없어 메일을 발송하지 않습니다.")
		return
	}

	var names []string
	for _, bp := range birthdayPeople {
		names = append(names, fmt.Sprintf("%s %s님", bp.Name, bp.Position))
	}

	subject := strings.Join(names, ", ") + "의 생일을 축하합니다~"
	todayStr := time.Now().Format("1월 2일")

	body := "오늘(" + todayStr + ")은 " + strings.Join(names, ", ") + "의 생일입니다.\n" +
		"모두 함께 축하해 주세요!\n\n" +
		"항상 건강하고 승승장구하시길 기원합니다.\n" +
		"동료분들~ 오늘 하루 축하 / 생일빵(?) 마구마구 날려주세요~"

	var emailList []string
	for _, r := range toList {
		if r.Email != "" {
			emailList = append(emailList, r.Email)
		}
	}
	to := strings.Join(emailList, ";")

	script := `
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.To = "` + to + `"
$mail.Subject = "` + subject + `"
$mail.Body = "` + body + `"
$mail.Send()
`

	err := exec.Command("powershell", "-Command", script).Run()
	if err == nil {
		fmt.Println(time.Now().Format("15:04:05") + " - 메일 발송 완료")
	} else {
		fmt.Println(time.Now().Format("15:04:05") + " - 메일 발송 실패: " + err.Error())
	}
}
