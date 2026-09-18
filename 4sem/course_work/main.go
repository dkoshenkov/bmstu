package main

import (
	"github.com/goriiin/cw/config"
	"github.com/goriiin/cw/pages"
	"github.com/goriiin/cw/storage"
	"html/template"
	"log"
	"net/http"
	"os"
)

func main() {
	cfg := config.MustLoad()
	s, err := storage.New(cfg.DbConfigPath)

	if err != nil {
		log.Fatalf("db error : %s", err)
		os.Exit(1)
	}
	err = s.TestSelect()
	if err != nil {
		log.Fatalf("db error : %s", err)
		os.Exit(1)
	}

	t, err := template.New("base").ParseFiles("./templates/index.gohtml", "./templates/table.gohtml")
	if err != nil {
		log.Print(err)
	}

	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		page := pages.InitBasePage()
		err := t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Printf("template error : %s", err)
		}
	})

	http.HandleFunc("/human", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitHuman(s)
		if err != nil {
			log.Printf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Printf("template error : %s", err)
		}
	})

	http.HandleFunc("/bankdata", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitBankData(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/cheque", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitCheque(s)
		if err != nil {
			log.Printf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Printf("template error : %s", err)
		}
	})

	http.HandleFunc("/damage", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitDamage(s)
		if err != nil {
			log.Printf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Printf("template error : %s", err)
		}
	})

	http.HandleFunc("/driverlicense", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitDriverLicense(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/fine", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitFine(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/generalinfo", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitGeneralInfo(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/regaddress", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitRegistrationAddress(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/rent", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitRent(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/transport", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitTransport(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/transportdamage", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitTransportDamages(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/transportinfo", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitTransportInfo(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/user", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitUser(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/userfines", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitUserFines(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.HandleFunc("/userlicense", func(w http.ResponseWriter, r *http.Request) {
		page, err := pages.InitUserLicense(s)
		if err != nil {
			log.Fatalf("db error : %s", err)
		}
		err = t.ExecuteTemplate(w, "index.gohtml", page)
		if err != nil {
			log.Fatalf("template error : %s", err)
		}
	})

	http.Handle("/static/", http.StripPrefix("/static/", http.FileServer(http.Dir("./static"))))

	_ = http.ListenAndServe(":8080", nil)
}
