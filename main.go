package main

import (
        "fmt"
        "net/http"
        "html/template"
        "io"
	      "os"
        "github.com/xuri/excelize/v2"
        "encoding/json"
        "strings"
        "io/ioutil"
        )


type Subject struct {
  NumberOfSubject string `json:"numberOfSubject"`
  TypeOfSubject string `json:"typeOfSubject"`
  NameDay string `json:"-"`
  SubjectName string `json:"subjectName"`
  Teacher string `json:"teacher"`
  Location string `json:"location"`
}

type Weekdays struct {
  NameGroup string
  Monday []Subject
  Tuesday []Subject
  Wednesday []Subject
  Thursday []Subject
  Friday []Subject
  Saturday []Subject
}

type User struct {
	Name string `json:"name"`
	Role string `json:"role"`
}

func ParseExcelFile(Group string, NumberOfWeek int) {
    f, err := excelize.OpenFile("ScheduleExcel.xlsx")
    if err != nil {
        fmt.Println(err)
        return
    }

    rows, err := f.GetRows("курс 1 ПИ ")
    if err != nil {
        fmt.Println(err)
        return
    }

    indexWeek := 0
    NameOFfile := ""

    if NumberOfWeek == 1 {
      indexWeek = 0
      NameOFfile = "FirstWeekSchedule.json"
    } else if NumberOfWeek == 2 {
      indexWeek = 10
      NameOFfile = "SecondWeekSchedule.json"
    } else {
      return
    }

    var Lessons []Subject
    var infday []string
    WeekDays := Weekdays{}
    WeekDays.NameGroup = Group
    NowDay := ""
    IndexOFGroup := 0

    OuterLoop:
    for index, row := range rows {
      if index == 2 {
        for ind, gp := range row {
          if ind > indexWeek {
            if gp == Group {
              IndexOFGroup = ind
              break OuterLoop
            }
          }
        }
      }
    }

    for index, row := range rows {

      if index > 3 {
        if len(row) > 3 {
          lesson := Subject{}

          if len(row[0]) != 0 {
            NowDay = row[0]
          }

          if len(row) > IndexOFGroup {
            if len(row[IndexOFGroup]) != 0 {
              if len(infday) == 0 {
                infday = append(infday, row[1], row[IndexOFGroup - 1], NowDay)
              }
              infday = append(infday, row[IndexOFGroup])
            }
          }
          if len(infday) == 6 {
              lesson.NumberOfSubject = infday[0]
              lesson.TypeOfSubject = infday[1]
              lesson.NameDay = infday[2]
              lesson.SubjectName = infday[3]
              lesson.Teacher = infday[4]
              lesson.Location = infday[5]

              Lessons = append(Lessons, lesson)
              infday = nil
          }
        }
      }
    }

    for  _, Day := range Lessons {
      switch Day.NameDay {
      case "понедельник":
        WeekDays.Monday = append(WeekDays.Monday, Day)
      case "Вторник":
        WeekDays.Tuesday = append(WeekDays.Tuesday, Day)
      case "Среда":
        WeekDays.Wednesday = append(WeekDays.Wednesday, Day)
      case "Четверг":
        WeekDays.Thursday = append(WeekDays.Thursday, Day)
      case "Пятница":
        WeekDays.Friday = append(WeekDays.Friday, Day)
      case "Суббота":
        WeekDays.Saturday = append(WeekDays.Saturday, Day)
      }
    }

    ScheduleJson := Weekdays{NameGroup: WeekDays.NameGroup, Monday: WeekDays.Monday, Tuesday: WeekDays.Tuesday, Wednesday: WeekDays.Wednesday, Thursday: WeekDays.Thursday, Friday: WeekDays.Friday, Saturday: WeekDays.Saturday}

    jsonLessons, err := json.Marshal(ScheduleJson)
    if err != nil {
        fmt.Println(err)
        return
    }

    file, err := os.OpenFile(NameOFfile, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
    if err != nil {
        fmt.Println(err)
        return
    }
    defer file.Close()

    file.Write(jsonLessons)
}

func Parse_File_To_Json() {

  f, err := excelize.OpenFile("ScheduleExcel.xlsx")
  if err != nil {
      fmt.Println(err)
      return
  }

  rows, err := f.GetRows("курс 1 ПИ ")
  if err != nil {
      fmt.Println(err)
      return
  }

  var gpp []string
  for i := 1; i < 3; i++ {
    OuterLoop:
    for index, row := range rows {
      if index == 2 {
        for ind, gp := range row {
          if ind > 2 {
            if len(gp) > 0 && strings.Contains(gp, "группа")  {
                ParseExcelFile(gp, i)
                gpp = append(gpp, gp)

              }
              if len(gpp) == 3 {
                gpp = nil
                break OuterLoop
              }
            }
          }
        }
      }
    }
  }

func index(w http.ResponseWriter, r *http.Request) {
  t, err := template.ParseFiles("templates/home_page.html")
  if err != nil {
    fmt.Fprintf(w, err.Error())
  }

  t.Execute(w, nil)
}



func schedule(w http.ResponseWriter, r *http.Request) {
  t, err := template.ParseFiles("templates/schedule.html")
  if err != nil {
    fmt.Fprintf(w, err.Error())
  }

  t.Execute(w, nil)
}

func uploadFile(w http.ResponseWriter, r *http.Request) {
	err := r.ParseMultipartForm(10 << 20)
	if err != nil {
		http.Error(w, "Error parsing the multipart form", http.StatusInternalServerError)
		return
	}

	file, _, err := r.FormFile("excelFile")
	if err != nil {
		http.Error(w, "Error retrieving the file from form data", http.StatusBadRequest)
		return
	}
	defer file.Close()

	out, err := os.OpenFile("ScheduleExcel.xlsx", os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0666)
	if err != nil {
		http.Error(w, "Error opening the file", http.StatusInternalServerError)
		return
	}
	defer out.Close()

	_, err = io.Copy(out, file)
	if err != nil {
		http.Error(w, "Error writing to the file", http.StatusInternalServerError)
		return
	}

	fmt.Fprintf(w, "File successfully received and saved as ScheduleExcel.xlsx.")

  Parse_File_To_Json()

}


func getUsersFromJSON(filename string) ([]User, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	bytes, err := ioutil.ReadAll(file)
	if err != nil {
		return nil, err
	}

	var users []User
	if err := json.Unmarshal(bytes, &users); err != nil {
		return nil, err
	}

	return users, nil
}

func updateUserRole(name string, role string) {
  users, err := getUsersFromJSON("users.json")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	for i, user := range users {
		if user.Name == name {
			users[i].Role = role
		}
	}
	file, _ := os.Create("users.json")
	defer file.Close()
	json.NewEncoder(file).Encode(users)
}

func users(w http.ResponseWriter, r *http.Request) {
  t, err := template.ParseFiles("templates/user_page.html")
  if err != nil {
    fmt.Fprintf(w, err.Error())
  }

  if r.Method == http.MethodGet {
    users, err := getUsersFromJSON("users.json")
  	if err != nil {
  		fmt.Println("Error:", err)
  		return
  	}
		t.Execute(w, users)
	} else if r.Method == http.MethodPost {
		r.ParseForm()
		name := r.Form.Get("name")
		role := r.Form.Get("role")
		updateUserRole(name, role)
		http.Redirect(w, r, "/users", http.StatusSeeOther)
	}
}

func handleFunc() {
  http.HandleFunc("/", index)
  http.HandleFunc("/user", users)
  http.HandleFunc("/schedule", schedule)
  http.HandleFunc("/upload", uploadFile)
  http.ListenAndServe(":8080", nil)
}

func main() {
  handleFunc()
}
