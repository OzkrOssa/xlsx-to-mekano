package main

import (
	"fmt"
	"log"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/mozillazg/go-unidecode"
	"github.com/xuri/excelize/v2"
)

type MekanoDataStruct struct {
	Tipo          string
	Prefijo       string
	Numero        string
	Secuencia     string
	Fecha         string
	Cuenta        string
	Terceros      string
	CentroCostos  string
	Nota          string
	Debito        string
	Credito       string
	Base          string
	Aplica        string
	TipoAnexo     string
	PrefijoAnexo  string
	NumeroAnexo   string
	Usuario       string
	Signo         string
	CuentaCobrar  string
	CuentaPagar   string
	NombreTercero string
	NombreCentro  string
	Interface     string
}

type MekanoInterface interface {
	Payment(file string) (string, error)
	Billing(file string, extras string) (string, error)
}

type mekanoRepository struct {
	dr         DatabaseInterface
	statistics StatisticInterface
}

func NewMekano(dr DatabaseInterface) MekanoInterface {
	sta := NewStatistics(dr)
	return &mekanoRepository{
		dr,
		sta,
	}
}

func (mr *mekanoRepository) Payment(file string) (string, error) {
	mr.statistics.SetFile(file)
	var paymentDataSlice []MekanoDataStruct
	var consecutive, rowCount int = 0, 0

	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return "", err
	}

	excelRows, err := xlsx.GetRows(xlsx.GetSheetName(0))
	if err != nil {
		fmt.Println(err)
		return "", err
	}

	c, err := mr.dr.GetPayment()
	if err != nil {
		return "", err
	}

	for _, row := range excelRows[1:] {
		cashier, err := mr.dr.GetCashiers(row[9])
		if err != nil {
			return "", err
		}
		rowCount++
		consecutive = c.Consecutive + rowCount

		paymentData := MekanoDataStruct{
			Tipo:          "RC",
			Prefijo:       "_",
			Numero:        strconv.Itoa(consecutive),
			Secuencia:     "",
			Fecha:         row[4],
			Cuenta:        "13050501",
			Terceros:      row[1],
			CentroCostos:  "C1",
			Nota:          "RECAUDO POR VENTA SERVICIOS",
			Debito:        "0",
			Credito:       row[5],
			Base:          "0",
			Aplica:        "",
			TipoAnexo:     "",
			PrefijoAnexo:  "",
			NumeroAnexo:   "",
			Usuario:       "SUPERVISOR",
			Signo:         "",
			CuentaCobrar:  "",
			CuentaPagar:   "",
			NombreTercero: row[2],
			NombreCentro:  "CENTRO DE COSTOS GENERAL",
			Interface:     time.Now().Format("02/01/2006 15:04"),
		}
		paymentDataSlice = append(paymentDataSlice, paymentData)

		paymentData2 := MekanoDataStruct{
			Tipo:          "RC",
			Prefijo:       "_",
			Numero:        strconv.Itoa(consecutive),
			Secuencia:     "",
			Fecha:         row[4],
			Cuenta:        cashier.Code,
			Terceros:      row[1],
			CentroCostos:  "C1",
			Nota:          "RECAUDO POR VENTA SERVICIOS",
			Debito:        row[5],
			Credito:       "0",
			Base:          "0",
			Aplica:        "",
			TipoAnexo:     "",
			PrefijoAnexo:  "",
			NumeroAnexo:   "",
			Usuario:       "SUPERVISOR",
			Signo:         "",
			CuentaCobrar:  "",
			CuentaPagar:   "",
			NombreTercero: row[2],
			NombreCentro:  "CENTRO DE COSTOS GENERAL",
			Interface:     time.Now().Format("02/01/2006 15:04"),
		}
		paymentDataSlice = append(paymentDataSlice, paymentData2)
	}
	exporterFile(paymentDataSlice)
	data := mr.statistics.Payment(paymentDataSlice, c.Consecutive, consecutive)
	return data, nil
}

func (mr *mekanoRepository) Billing(file string, extras string) (string, error) {

	var montoBaseFinal float64
	var montoIvaFinal float64
	var montoDebitoFinal float64
	var itemIvaBaseFinal float64

	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return "", err
	}

	billingFile, err := xlsx.GetRows(xlsx.GetSheetName(0))
	if err != nil {
		fmt.Println(err)
		return "", err
	}

	ivaXlsx, err := excelize.OpenFile(extras)
	if err != nil {
		fmt.Println(err)
		return "", err
	}

	itemsIvaFile, err := ivaXlsx.GetRows(ivaXlsx.GetSheetName(0))
	if err != nil {
		log.Println(err, "itemsIvaFile")
		return "", err
	}

	var BillingDataSheet []MekanoDataStruct

	for _, bRow := range billingFile[1:] {

		montoDebito, err := strconv.ParseFloat(bRow[14], 64)
		if err != nil {
			log.Println(err, "MontoDebito")
		}
		_, decimalDebito := math.Modf(montoDebito)
		if decimalDebito >= 0.5 {
			montoDebitoFinal = math.Ceil(montoDebito)
		} else {
			montoDebitoFinal = math.Round(montoDebito)
		}

		montoBase, err := strconv.ParseFloat(bRow[12], 64)
		if err != nil {
			log.Println(err, "MontoBase")
		}
		_, decimalBase := math.Modf(montoBase)
		if decimalBase >= 0.5 {
			montoBaseFinal = math.Ceil(montoBase)
		} else {
			montoBaseFinal = math.Round(montoBase)
		}

		montoIva, err := strconv.ParseFloat(strings.TrimSpace(bRow[13]), 64)
		if err != nil {
			log.Println(err, "MontoIva")
		}
		_, decimalIva := math.Modf(montoIva)

		if decimalIva >= 0.5 {
			montoIvaFinal = math.Ceil(montoIva)
		} else {
			montoIvaFinal = math.Round(montoIva)
		}

		if !strings.Contains(bRow[21], ",") {
			account, err := mr.dr.GetAccounts(bRow[21])
			if err != nil {
				return "", fmt.Errorf("error to get account %s", bRow[21])
			}

			costCenter, err := mr.dr.GetCostCenter(unidecode.Unidecode(bRow[17]))
			if err != nil {
				return "", fmt.Errorf("error to get cost center %s", bRow[17])
			}
			billingNormal := MekanoDataStruct{
				Tipo:          "FVE",
				Prefijo:       "_",
				Numero:        bRow[8],
				Secuencia:     "",
				Fecha:         bRow[9],
				Cuenta:        account.Code,
				Terceros:      bRow[1],
				CentroCostos:  costCenter.Code,
				Nota:          "FACTURA ELECTRÓNICA DE VENTA",
				Debito:        "0",
				Credito:       fmt.Sprintf("%f", montoBaseFinal),
				Base:          "0",
				Aplica:        "",
				TipoAnexo:     "",
				PrefijoAnexo:  "",
				NumeroAnexo:   "",
				Usuario:       "SUPERVISOR",
				Signo:         "",
				CuentaCobrar:  "",
				CuentaPagar:   "",
				NombreTercero: bRow[2],
				NombreCentro:  bRow[17],
				Interface:     time.Now().Format("02/01/2006 15:04"),
			}

			BillingDataSheet = append(BillingDataSheet, billingNormal)

			billingIva := MekanoDataStruct{
				Tipo:          "FVE",
				Prefijo:       "_",
				Numero:        bRow[8],
				Secuencia:     "",
				Fecha:         bRow[9],
				Cuenta:        "24080505",
				Terceros:      bRow[1],
				CentroCostos:  costCenter.Code,
				Nota:          "FACTURA ELECTRÓNICA DE VENTA",
				Debito:        "0",
				Credito:       fmt.Sprintf("%f", montoIvaFinal),
				Base:          fmt.Sprintf("%f", montoBaseFinal),
				Aplica:        "",
				TipoAnexo:     "",
				PrefijoAnexo:  "",
				NumeroAnexo:   "",
				Usuario:       "SUPERVISOR",
				Signo:         "",
				CuentaCobrar:  "",
				CuentaPagar:   "",
				NombreTercero: bRow[2],
				NombreCentro:  bRow[17],
				Interface:     time.Now().Format("02/01/2006 15:04"),
			}

			BillingDataSheet = append(BillingDataSheet, billingIva)

			billingCxC := MekanoDataStruct{
				Tipo:          "FVE",
				Prefijo:       "_",
				Numero:        bRow[8],
				Secuencia:     "",
				Fecha:         bRow[9],
				Cuenta:        "13050501",
				Terceros:      bRow[1],
				CentroCostos:  costCenter.Code,
				Nota:          "FACTURA ELECTRÓNICA DE VENTA",
				Debito:        fmt.Sprintf("%f", montoDebitoFinal),
				Credito:       "0",
				Base:          "0",
				Aplica:        "",
				TipoAnexo:     "",
				PrefijoAnexo:  "",
				NumeroAnexo:   "",
				Usuario:       "SUPERVISOR",
				Signo:         "",
				CuentaCobrar:  "",
				CuentaPagar:   "",
				NombreTercero: bRow[2],
				NombreCentro:  bRow[17],
				Interface:     time.Now().Format("02/01/2006 15:04"),
			}

			BillingDataSheet = append(BillingDataSheet, billingCxC)
		} else {
			costCenter, err := mr.dr.GetCostCenter(unidecode.Unidecode(bRow[17]))
			if err != nil {
				err2 := fmt.Errorf("error to get cost center %s", bRow[17])
				return "", err2
			}
			splitBillingItems := strings.Split(bRow[21], ",")
			for _, item := range splitBillingItems {
				for _, itemIva := range itemsIvaFile[1:] {

					if itemIva[1] == strings.TrimSpace(item) && itemIva[0] == bRow[0] {
						itemIvaBase, _ := strconv.ParseFloat(itemIva[2], 64)
						_, decimalIvaBase := math.Modf(itemIvaBase)

						if decimalIvaBase >= 0.5 {
							itemIvaBaseFinal = math.Ceil(itemIvaBase)
						} else {
							itemIvaBaseFinal = math.Round(itemIvaBase)
						}

						account, err := mr.dr.GetAccounts(unidecode.Unidecode(strings.TrimSpace(item)))
						if err != nil {
							err2 := fmt.Errorf("error to get account %s", strings.TrimSpace(item))
							return "", err2

						}

						costCenter, err := mr.dr.GetCostCenter(unidecode.Unidecode(bRow[17]))
						if err != nil {
							err2 := fmt.Errorf("error to get account %s", unidecode.Unidecode(bRow[17]))
							return "", err2
						}

						billingNormalPlus := MekanoDataStruct{
							Tipo:          "FVE",
							Prefijo:       "_",
							Numero:        bRow[8],
							Secuencia:     "",
							Fecha:         bRow[9],
							Cuenta:        account.Code,
							Terceros:      bRow[1],
							CentroCostos:  costCenter.Code,
							Nota:          "FACTURA ELECTRÓNICA DE VENTA",
							Debito:        "0",
							Credito:       fmt.Sprintf("%f", itemIvaBaseFinal),
							Base:          "0",
							Aplica:        "",
							TipoAnexo:     "",
							PrefijoAnexo:  "",
							NumeroAnexo:   "",
							Usuario:       "SUPERVISOR",
							Signo:         "",
							CuentaCobrar:  "",
							CuentaPagar:   "",
							NombreTercero: bRow[2],
							NombreCentro:  bRow[17],
							Interface:     time.Now().Format("02/01/2006 15:04"),
						}
						BillingDataSheet = append(BillingDataSheet, billingNormalPlus)
					}
				}
			}
			billingIvaPlus := MekanoDataStruct{
				Tipo:          "FVE",
				Prefijo:       "_",
				Numero:        bRow[8],
				Secuencia:     "",
				Fecha:         bRow[9],
				Cuenta:        "24080505",
				Terceros:      bRow[1],
				CentroCostos:  costCenter.Code,
				Nota:          "FACTURA ELECTRÓNICA DE VENTA",
				Debito:        "0",
				Credito:       fmt.Sprintf("%f", montoIvaFinal),
				Base:          fmt.Sprintf("%f", montoBaseFinal),
				Aplica:        "",
				TipoAnexo:     "",
				PrefijoAnexo:  "",
				NumeroAnexo:   "",
				Usuario:       "SUPERVISOR",
				Signo:         "",
				CuentaCobrar:  "",
				CuentaPagar:   "",
				NombreTercero: bRow[2],
				NombreCentro:  bRow[17],
				Interface:     time.Now().Format("02/01/2006 15:04"),
			}

			BillingDataSheet = append(BillingDataSheet, billingIvaPlus)

			billingCxCPlus := MekanoDataStruct{
				Tipo:          "FVE",
				Prefijo:       "_",
				Numero:        bRow[8],
				Secuencia:     "",
				Fecha:         bRow[9],
				Cuenta:        "13050501",
				Terceros:      bRow[1],
				CentroCostos:  costCenter.Code,
				Nota:          "FACTURA ELECTRÓNICA DE VENTA",
				Debito:        fmt.Sprintf("%f", montoDebitoFinal),
				Credito:       "0",
				Base:          "0",
				Aplica:        "",
				TipoAnexo:     "",
				PrefijoAnexo:  "",
				NumeroAnexo:   "",
				Usuario:       "SUPERVISOR",
				Signo:         "",
				CuentaCobrar:  "",
				CuentaPagar:   "",
				NombreTercero: bRow[2],
				NombreCentro:  bRow[17],
				Interface:     time.Now().Format("02/01/2006 15:04"),
			}

			BillingDataSheet = append(BillingDataSheet, billingCxCPlus)
		}
	}

	exporterFile(BillingDataSheet)
	data := mr.statistics.Billing(BillingDataSheet)
	return data, nil
}

func exporterFile(mekanoData []MekanoDataStruct) {
	f := excelize.NewFile()
	// Crea un nuevo sheet.
	index, _ := f.NewSheet("Sheet1")

	// Define los títulos de las columnas
	f.SetCellValue("Sheet1", "A1", "TIPO")
	f.SetCellValue("Sheet1", "B1", "PREFIJO")
	f.SetCellValue("Sheet1", "C1", "NUMERO")
	f.SetCellValue("Sheet1", "D1", "FECHA")
	f.SetCellValue("Sheet1", "E1", "CUENTA")
	f.SetCellValue("Sheet1", "F1", "TERCERO")
	f.SetCellValue("Sheet1", "G1", "CENTRO")
	f.SetCellValue("Sheet1", "H1", "DETALLE")
	f.SetCellValue("Sheet1", "I1", "DEBITO")
	f.SetCellValue("Sheet1", "J1", "CREDITO")
	f.SetCellValue("Sheet1", "K1", "BASE")
	f.SetCellValue("Sheet1", "L1", "USUARIO")
	f.SetCellValue("Sheet1", "M1", "NOMBRE TERCERO")
	f.SetCellValue("Sheet1", "N1", "NOMBRE CENTRO")

	// Poblar datos
	for i, m := range mekanoData {
		row := i + 2 // Comenzar en la fila 2
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), m.Tipo)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), m.Prefijo)
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), m.Numero)
		f.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), m.Fecha)
		f.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), m.Cuenta)
		f.SetCellValue("Sheet1", fmt.Sprintf("F%d", row), m.Terceros)
		f.SetCellValue("Sheet1", fmt.Sprintf("G%d", row), m.CentroCostos)
		f.SetCellValue("Sheet1", fmt.Sprintf("H%d", row), m.Nota)
		f.SetCellValue("Sheet1", fmt.Sprintf("I%d", row), m.Debito)
		f.SetCellValue("Sheet1", fmt.Sprintf("J%d", row), m.Credito)
		f.SetCellValue("Sheet1", fmt.Sprintf("K%d", row), m.Base)
		f.SetCellValue("Sheet1", fmt.Sprintf("L%d", row), m.Usuario)
		f.SetCellValue("Sheet1", fmt.Sprintf("M%d", row), m.NombreTercero)
		f.SetCellValue("Sheet1", fmt.Sprintf("N%d", row), m.NombreCentro)
	}

	// Establece el sheet activo al primero
	f.SetActiveSheet(index)

	dir, err := os.UserHomeDir()
	if err != nil {
		fmt.Println(err)

	}

	// Guarda el archivo de Excel
	if err := f.SaveAs(filepath.Join(dir, "CONTABLE.xlsx")); err != nil {
		fmt.Println(err)
	}
}
