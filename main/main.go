package main

import (
	"context"
	"fmt"
	"github.com/bwmarrin/discordgo"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"log"
	"os"
	"strconv"
	"strings"
	"time"
)

const DateFormat = "1/2/2006 15:04:05"

const Role = "1113654375135584296"

const GuildID = "543724459476123669"

const ResponseRange = "A2:F"

// SpreadSheetID https://docs.google.com/spreadsheets/d/10SAgHoWG5hbwESDMLvuDGANOWK3f-b5VoQ-Y-XQh5jg/edit?resourcekey#gid=590955473
const SpreadSheetID = "10SAgHoWG5hbwESDMLvuDGANOWK3f-b5VoQ-Y-XQh5jg"
const SpreadSheetGID = 590955473

type FormResponse struct {
	Time        *time.Time
	RowIndex    int64
	NetID       string
	Name        string
	DiscordUser string
	Attendance  uint
}

func main() {

	ctx := context.Background()

	discordToken, err := os.ReadFile("discord-token.txt")
	if err != nil {
		log.Fatalf("Unable to read discord token: %v", err)
	}

	discord, err := discordgo.New("Bot " + string(discordToken))
	if err != nil {
		log.Fatalf("Unable to create discord session: %v", err)
	}

	sheetsService, err := sheets.NewService(ctx, option.WithCredentialsFile("sheets-api-key.json"))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	fmt.Println("https://discord.com/oauth2/authorize?client_id=1113609197049892884&scope=bot&permissions=275146411008")

	// Start checking for updates forever
	for {
		Update(discord, sheetsService)
		time.Sleep(1 * time.Minute)
	}

}

func Update(discord *discordgo.Session, sheetsService *sheets.Service) {

	fmt.Println("Getting Responses")
	responses := GetResponses(sheetsService)

	fmt.Println("Updating Discord Members")
	UpdateDiscordMembers(responses, discord)

	fmt.Println("Updating Form")
	UpdateForm(responses, sheetsService)

	fmt.Println("Done")
	fmt.Println()
}

func UpdateDiscordMembers(formResponses []FormResponse, discord *discordgo.Session) {

	formResponseByDiscordUser := map[string]FormResponse{}
	for _, formResponse := range formResponses {
		formResponseByDiscordUser[formResponse.DiscordUser] = formResponse
	}

	members, err := discord.GuildMembers(GuildID, "", 1000)
	if err != nil {
		log.Fatalf("Unable to retrieve guild members: %v", err)
	}

	for _, member := range members {

		_, ok := formResponseByDiscordUser[member.User.Username+"#"+member.User.Discriminator]
		if !ok {
			continue
		}

		hasRole := false
		for _, role := range member.Roles {
			if role == Role {
				hasRole = true
				break
			}
		}

		if !hasRole {
			if err := discord.GuildMemberRoleAdd(GuildID, member.User.ID, Role); err != nil {
				fmt.Println("Unable to add role to user:", member.User.ID, err)
				continue
			}
		}
	}
}

func UpdateForm(formResponses []FormResponse, sheetsService *sheets.Service) {

	dayFormat := "2006-01-02"

	// NetID + day
	netIDAndDay := map[string]struct{}{}

	// Build a map to count attendance by NetID
	netIDAttendanceCount := make(map[string]uint)
	for _, response := range formResponses {

		day := response.Time.Format(dayFormat)
		key := response.NetID + day

		if _, ok := netIDAndDay[key]; !ok {
			netIDAttendanceCount[response.NetID] += response.Attendance
		}

		netIDAndDay[key] = struct{}{}
	}

	var requests []*sheets.Request

	for _, response := range formResponses {

		newAttendance, ok := netIDAttendanceCount[response.NetID]
		if !ok || response.Attendance == newAttendance {
			continue
		}

		response.Attendance = netIDAttendanceCount[response.NetID]
		attendance := float64(response.Attendance)

		requests = append(requests, &sheets.Request{
			UpdateCells: &sheets.UpdateCellsRequest{
				Start: &sheets.GridCoordinate{
					SheetId:     SpreadSheetGID, // Your sheet ID
					RowIndex:    response.RowIndex + 1,
					ColumnIndex: 4, // Column E
				},
				Rows: []*sheets.RowData{
					{
						Values: []*sheets.CellData{
							{
								UserEnteredValue: &sheets.ExtendedValue{NumberValue: &attendance},
							},
						},
					},
				},
				Fields: "*",
			},
		})

		currentDate := time.Now().Format(DateFormat)

		requests = append(requests, &sheets.Request{
			UpdateCells: &sheets.UpdateCellsRequest{
				Start: &sheets.GridCoordinate{
					SheetId:     SpreadSheetGID, // Your sheet ID
					RowIndex:    response.RowIndex + 1,
					ColumnIndex: 0, // Column G
				},
				Rows: []*sheets.RowData{
					{
						Values: []*sheets.CellData{
							{
								UserEnteredValue: &sheets.ExtendedValue{StringValue: &currentDate},
							},
						},
					},
				},
				Fields: "*",
			},
		})
	}

	deleteDuplicates := sheets.DeleteDuplicatesRequest{
		Range: &sheets.GridRange{
			SheetId:       SpreadSheetGID,
			StartRowIndex: 1,
		},
		ComparisonColumns: []*sheets.DimensionRange{
			{
				SheetId:    SpreadSheetGID,
				Dimension:  "COLUMNS",
				StartIndex: 1, // NetID is in the 1st column (zero-indexed)
				EndIndex:   2, // The next column after NetID
			},
		},
	}

	_, err := sheetsService.Spreadsheets.BatchUpdate(SpreadSheetID, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: append(requests, &sheets.Request{DeleteDuplicates: &deleteDuplicates}),
	}).Do()
	if err != nil {
		log.Fatalf("[1] Unable to update attendance: %v", err)
	}

}

func GetResponses(sheetsService *sheets.Service) []FormResponse {

	resp, err := sheetsService.Spreadsheets.Values.Get(SpreadSheetID, ResponseRange).Do()
	if err != nil {
		log.Fatalf("Unable to retrieve data from sheet: %v", err)
	}

	if len(resp.Values) == 0 {
		fmt.Println("No data found.")
		return nil
	}

	responses := make([]FormResponse, 0, len(resp.Values))
	for index, row := range resp.Values {

		var values []string

		for _, value := range row {

			valueAsString, ok := value.(string)
			if !ok || strings.TrimSpace(valueAsString) == "" {
				continue
			}

			values = append(values, valueAsString)
		}

		if len(values) == 0 {
			continue
		}

		if len(values) < 5 {
			fmt.Println("Unable to parse values:", row)
			continue
		}

		attendance := uint64(1)

		if len(values) > 5 {
			attendance, err = strconv.ParseUint(values[4], 10, 32)
			if err != nil {
				fmt.Println("Failed to parse attendance:", row, err)
				continue
			}
		}

		responseDate, err := time.Parse(DateFormat, values[0])
		if err != nil {
			fmt.Println("Unable to parse time:", row, err)
			continue
		}

		responses = append(responses, FormResponse{
			Time:        &responseDate,
			RowIndex:    int64(index),
			NetID:       values[1],
			Name:        values[2],
			DiscordUser: values[3],
			Attendance:  uint(attendance),
		})
	}

	return responses
}
