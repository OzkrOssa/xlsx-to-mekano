package main

import (
	"encoding/json"
	"fmt"
	"log"
	"strconv"
	"time"
)

type paymentStatistics struct {
	FileName    string `json:"archivo"`
	RangoRC     string `json:"rango_rc"`
	Bancolombia int    `json:"bancolombia"`
	Davivienda  int    `json:"davivienda"`
	Susuerte    int    `json:"susuerte"`
	PayU        int    `json:"payu"`
	Efectivo    int    `json:"efectivo"`
	Total       int    `json:"total"`
}
type billingStatistics struct {
	File    string  `json:"file"`
	Debito  float64 `json:"debito"`
	Credito float64 `json:"credito"`
	Base    float64 `json:"base"`
}

type StatisticInterface interface {
	SetFile(filename string)
	Payment(mekanoData []MekanoDataStruct, initRange int, lastRange int) string
	Billing(mekanoData []MekanoDataStruct) string
}

type statistics struct {
	db   DatabaseInterface
	file string
}

func NewStatistics(db DatabaseInterface) StatisticInterface {
	return &statistics{
		db: db,
	}
}

func (s *statistics) SetFile(filename string) {
	s.file = filename
}

func (s statistics) Payment(mekanoData []MekanoDataStruct, initRange int, lastRange int) string {
	var efectivo, bancolombia, davivienda, susuerte, payU, total int = 0, 0, 0, 0, 0, 0

	for _, d := range mekanoData {
		debito, err := strconv.Atoi(d.Debito)
		total += debito
		if err != nil {
			log.Println(err)
		}
		switch d.Cuenta {
		case "11050501": //Efectivo
			efectivo += debito
		case "11200501": //Bancolombia
			bancolombia += debito
		case "11200510": //Davivienda
			davivienda += debito
		case "13452505": //Susuerte
			susuerte += debito
		case "13452501": //Pay U
			payU += debito
		}
	}

	var statistic = paymentStatistics{
		FileName:    s.file,
		RangoRC:     fmt.Sprintf("%d-%d", initRange+1, lastRange),
		Efectivo:    efectivo,
		Bancolombia: bancolombia,
		Davivienda:  davivienda,
		PayU:        payU,
		Susuerte:    susuerte,
		Total:       total,
	}
	result, err := json.Marshal(statistic)
	if err != nil {
		log.Println(err)
	}

	_, err = s.db.SavePayment(Payment{Consecutive: lastRange, CreateAt: time.Now(), File: s.file})
	if err != nil {
		log.Println(err)
	}
	return string(result)
}

func (s statistics) Billing(mekanoData []MekanoDataStruct) string {
	var d, c, b float64 = 0, 0, 0
	for _, row := range mekanoData {
		debito, _ := strconv.ParseFloat(row.Debito, 64)
		d += debito
		credito, _ := strconv.ParseFloat(row.Credito, 64)
		c += credito
		base, _ := strconv.ParseFloat(row.Base, 64)
		b += base
	}

	bs := billingStatistics{
		File:    s.file,
		Debito:  d,
		Credito: c,
		Base:    b,
	}

	result, err := json.Marshal(bs)
	if err != nil {
		log.Println(err)
	}

	_, err = s.db.SaveBilling(Billing{File: s.file, Base: int(b), Debit: int(d), Credit: int(c), CreateAt: time.Now()})
	if err != nil {
		log.Println(err)
	}

	return string(result)
}
