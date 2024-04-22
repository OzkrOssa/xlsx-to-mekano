package main

func main() {
	db, err := NewDatabase()
	if err != nil {
		panic(err)
	}

	mekano := NewMekano(db)
	mekano.Payment("filepath")
}
