package main

import (
	"bufio"
	"fmt"
	"os"
	"strings"
	"time"
)

func main() {
	var recipients []Recipient
	var alarmTime string

	go func() {
		for {
			now := time.Now()
			if alarmTime != "" && FmtTime(now) == alarmTime && len(recipients) > 0 {
				today := now.Format("01-02")
				birthdayPeople := GetBirthdayPeople(today, recipients)
				SendMail(recipients, birthdayPeople)
				time.Sleep(61 * time.Second)
			}
			time.Sleep(1 * time.Second)
		}
	}()

	scanner := bufio.NewScanner(os.Stdin)
	for {
		var prompt string
		if len(recipients) > 0 {
			next := GetNextBirthdays(recipients, 3)
			if len(next) > 0 {
				var nameList []string
				for _, p := range next {
					dateStr := fmt.Sprintf("(%d월 %d일)", p.Month, p.Day)
					nameList = append(nameList, p.Name+" "+p.Position+" "+dateStr)
				}
				prompt = "[다음 생일자: " + strings.Join(nameList, ", ") + "] "
			}
		}
		fmt.Print(prompt + "> ")

		if !scanner.Scan() {
			break
		}
		cmd := strings.TrimSpace(scanner.Text())

		switch cmd {
		case "":
			// 아무 반응 없이 넘어감
			continue

		case "reload", "r":
			r, err := LoadRecipients("members.txt")
			if err != nil {
				fmt.Println("수신자 목록 불러오기 실패:", err)
				continue
			}
			recipients = r
			fmt.Println("수신자 목록이 갱신되었습니다.")

		case "send", "s":
			if len(recipients) == 0 {
				fmt.Println("수신자 목록이 비어 있습니다. 먼저 'reload'를 실행하세요.")
				continue
			}
			today := time.Now().Format("01-02")
			birthdayPeople := GetBirthdayPeople(today, recipients)
			SendMail(recipients, birthdayPeople)

		case "list", "l":
			if len(recipients) == 0 {
				fmt.Println("수신자 목록이 없습니다.")
				continue
			}
			PrintRecipients(recipients)

		case "alarm", "a":
			if alarmTime != "" {
				fmt.Printf("현재 알람 시간: %s시 %s분\n", alarmTime[:2], alarmTime[2:])
			} else {
				fmt.Println("현재 알람 시간: 설정되지 않음")
			}
			fmt.Print("새 알람 시간을 입력하세요 (예: 0600 → 오전 6시): ")
			if scanner.Scan() {
				input := strings.TrimSpace(scanner.Text())
				if len(input) == 4 && IsDigits(input) {
					alarmTime = input
					fmt.Printf("알람 시간이 %s시 %s분으로 설정되었습니다.\n", input[:2], input[2:])
				} else {
					fmt.Println("잘못된 형식입니다. 4자리 숫자(HHMM)를 입력하세요.")
				}
			}

		case "exit", "x", "quit", "q":
			return

		default:
			fmt.Println("알 수 없는 명령어입니다.")
		}

	}
}
