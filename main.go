package main

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
)

func main() {
	db, err := NewDatabase()
	if err != nil {
		panic(err)
	}

	a := app.New()
	w := a.NewWindow("xlsx to mekano")

	button := widget.NewButton("Subir Archivo", func() {
		dialog.ShowFileOpen(func(uc fyne.URIReadCloser, err error) {
			if err != nil {
				panic(err)
			}

			if uc == nil {
				return
			}

			defer uc.Close()

			mekano := NewMekano(db)
			mekano.Payment(uc.URI().Path())
		}, w)
	})

	w.SetContent(button)
	w.ShowAndRun()

}
