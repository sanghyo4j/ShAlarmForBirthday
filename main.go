package main

import (
	"bufio"
	"fmt"
	"os"
	"os/exec"
	"strconv"
	"strings"
	"time"
)

func promptAdvanceDays(days *int) {
	fmt.Print("며칠 전부터 생일을 챙기시겠습니까? (0 = 당일만): ")
	var input string
	fmt.Scanln(&input)
	d, err := strconv.Atoi(strings.TrimSpace(input))
	if err != nil || d < 0 {
		fmt.Println("잘못된 입력입니다. 기본값 0 사용")
		d = 0
	}
	*days = d
}

func main() {
	var recipients []Recipient
	var alarmTime string
	var advanceDays = 3

	cmd := exec.Command("cmd", "/c", "cls")
	cmd.Stdout = os.Stdout
	cmd.Run()

	go func() {
		for {
			if alarmTime != "" && FmtTime(time.Now()) == alarmTime && len(recipients) > 0 {
				birthdayPeople := GetBirthdaysWithinDays(recipients, advanceDays)
				if advanceDays == 0 {
					SendMailForToday(birthdayPeople, recipients)
				} else {
					SendMailForUpcoming(birthdayPeople, recipients)
				}
				time.Sleep(61 * time.Second)
			}
			time.Sleep(1 * time.Second)
		}
	}()

	scanner := bufio.NewScanner(os.Stdin)
	for {
		var prompt string
		if len(recipients) > 0 {
			next := GetBirthdaysWithinDays(recipients, advanceDays)
			if len(next) > 0 {
				var nameList []string
				for _, p := range next {
					dateStr := fmt.Sprintf("(%d월 %d일)", p.Month, p.Day)
					nameList = append(nameList, p.Name+" "+p.Position+" "+dateStr)
				}
				prompt = fmt.Sprintf("[다음 생일자(%d일 이내): %s] ", advanceDays, strings.Join(nameList, ", "))
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

		case "day", "d":
			promptAdvanceDays(&advanceDays)
			fmt.Printf("앞으로 %d일 이내 생일자에 대해 메일을 발송합니다.\n", advanceDays)

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

			fmt.Print("지금 바로 발송됩니다. 보낼까요? (y/n): ")
			if scanner.Scan() {
				confirm := strings.TrimSpace(scanner.Text())
				if strings.ToLower(confirm) != "y" {
					fmt.Println("발송이 취소되었습니다.")
					continue
				}
			}

			birthdayPeople := GetBirthdaysWithinDays(recipients, advanceDays)
			if advanceDays == 0 {
				SendMailForToday(birthdayPeople, recipients)
			} else {
				SendMailForUpcoming(birthdayPeople, recipients)
			}

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
