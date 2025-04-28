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

func GetBirthdaysWithinDays(list []Recipient, days int) []BirthdayPerson {
	today := time.Now()
	var result []BirthdayPerson

	for _, r := range list {
		if r.Birthday == "" {
			continue
		}
		parts := strings.Split(r.Birthday, "-")
		if len(parts) != 3 {
			continue
		}
		month, _ := strconv.Atoi(parts[1])
		day, _ := strconv.Atoi(parts[2])

		for d := 0; d <= days; d++ {
			checkDate := today.AddDate(0, 0, d)
			if int(checkDate.Month()) == month && checkDate.Day() == day {
				result = append(result, BirthdayPerson{
					Name:     r.Name,
					Position: r.Position,
					Month:    month,
					Day:      day,
				})
				break
			}
		}
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
	sort.Slice(list, func(i, j int) bool {
		// 생일이 없으면 뒤로 정렬
		if list[i].Birthday == "" {
			return false
		}
		if list[j].Birthday == "" {
			return true
		}

		partsI := strings.Split(list[i].Birthday, "-")
		partsJ := strings.Split(list[j].Birthday, "-")
		if len(partsI) < 3 || len(partsJ) < 3 {
			return false
		}

		monthI, _ := strconv.Atoi(partsI[1])
		dayI, _ := strconv.Atoi(partsI[2])
		monthJ, _ := strconv.Atoi(partsJ[1])
		dayJ, _ := strconv.Atoi(partsJ[2])

		if monthI == monthJ {
			return dayI < dayJ
		}
		return monthI < monthJ
	})

	fmt.Println("수신자 목록 (생일 월/일 순):")
	for _, r := range list {
		nameWithPosition := r.Name + " " + r.Position
		fmt.Printf(" - %-8s\t%s\t%s\n", nameWithPosition, r.Birthday, r.Email)
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

func SendMailForToday(birthdayPeople []BirthdayPerson, toList []Recipient) {
	if len(birthdayPeople) == 0 {
		fmt.Println("오늘 생일자가 없습니다. 메일을 발송하지 않습니다.")
		return
	}

	todayStr := time.Now().Format("1월 2일")

	var names []string
	for _, p := range birthdayPeople {
		names = append(names, p.Name+" "+p.Position+"님")
	}

	subject := strings.Join(names, ", ") + "의 생일을 축하합니다~"
	body := "오늘(" + todayStr + ")은 " + strings.Join(names, ", ") + "의 생일입니다.\n" +
		"모두 함께 축하해 주세요!\n\n" +
		"항상 건강하고 승승장구하시길 기원합니다.\n" +
		"동료분들~ 오늘 하루 축하 / 생일빵(?) 마구마구 날려주세요~"

	sendOutlookMail(toList, subject, body)
}

func SendMailForUpcoming(birthdayPeople []BirthdayPerson, toList []Recipient) {
	if len(birthdayPeople) == 0 {
		fmt.Println("예정된 생일자가 없습니다. 메일을 발송하지 않습니다.")
		return
	}

	var lines, names []string
	for _, p := range birthdayPeople {
		names = append(names, p.Name+" "+p.Position+"님")
		dateStr := fmt.Sprintf("%d월 %d일", p.Month, p.Day)
		lines = append(lines, dateStr+"은 "+p.Name+" "+p.Position+"님의 생일입니다.")
	}

	subject := strings.Join(names, ", ") + "의 생일을 축하합니다~"

	body := strings.Join(lines, "\n") + "\n" +
		"모두 함께 축하해 주세요!\n\n" +
		"항상 건강하고 승승장구하시길 기원합니다.\n" +
		"동료분들~ 오늘 하루 축하 / 생일빵(?) 마구마구 날려주세요~"

	sendOutlookMail(toList, subject, body)
}

func sendOutlookMail(toList []Recipient, subject, body string) {
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

$mail.Display()
Start-Sleep -Milliseconds 500

$inspector = $mail.GetInspector
$wordEditor = $inspector.WordEditor
$selection = $wordEditor.Application.Selection
$selection.HomeKey(6)
$selection.InlineShapes.AddPicture((Get-Item "` + "./happy_birthday.jpg" + `").FullName)
$selection.TypeParagraph()

$mail.Send()
`
	cmd := exec.Command("powershell", "-Command", script)
	output, err := cmd.CombinedOutput()

	if err == nil {
		fmt.Println(time.Now().Format("15:04:05") + " - 메일 발송 완료")
	} else {
		fmt.Println(time.Now().Format("15:04:05") + " - 메일 발송 실패")
		fmt.Println("PowerShell Error Output:\n" + string(output))
	}
}
