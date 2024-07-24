package main

import (
	"fmt"
	"log"
	"reflect"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Address struct {
	address1   string
	address2   string
	city       string
	state      string
	postalCode string
	country    string
	phone      string
	note       string
}

type Acrylic struct {
	bigKeero     int
	aline        int
	rhaine       int
	draven       int
	vespera      int
	gabe         int
	sash         int
	isaac        int
	felice       int
	theDarkQueen int
	spixx        int
	other        int
}

type Order struct {
	vol1SofCover            string
	extraVol2SofCover       string
	vol1ExtraBundle         string
	vol2SofCover            string
	doubleSidedBookmark     string
	vol2Booklet             string
	signedBookplate         string
	vol1SofCoverBook        string
	acrylicStandZevAviKeero string
	sketch                  string
	vol1ClockworkBook       string
	keeroKeychain           string
	vol1SignedBook          string
	acrylic                 Acrylic
}

type FullOrder struct {
	fullName              string
	email                 string
	backerUID             string
	address               Address
	order                 Order
	acrylicAddOnsNoFormat string
}

var orders []FullOrder

// Assuming following is the order of the headers
// Backer UID	Email	[Addon: 8250494] Volume 1 Softcover Book	[Addon: 8300565] Extra Copy of Volume 2 Softcover Book	[Addon: 8262164] Vol.1 Extras Bundle	Vol 2 Softcover Book	Double-Sided Bookmark	“The Clockwork of Everwake” Vol2 Booklet	Signed Bookplate	Vol 1 Softcover Book	Zevryx & Avi/Keero Acrylic Standee	Original Everwake Sketch ~ signed	Vol.1 Clockwork Booklet	Keero Keychain	Vol.1 Signed Bookplate	Shipping Name	Shipping Address 1	Shipping Address 2	Shipping City	Shipping State	Shipping Postal Code	Shipping Country Name	Shipping Phone Number	Shipping Delivery Notes	Acrylic Add Ons	n/a

func main() {
	f, err := excelize.OpenFile("./data/AllRewards.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	rows, err := f.Rows("rewards")
	if err != nil {
		log.Fatal(err)
	}

	for rows.Next() {

		row, err := rows.Columns()
		if err != nil {
			log.Fatal(err)
		}

		acrylics := acrylicSplit(row[24])

		newOrder := FullOrder{
			fullName:              row[15],
			email:                 row[1],
			backerUID:             row[0],
			acrylicAddOnsNoFormat: row[24],
			address: Address{
				address1:   row[16],
				address2:   row[17],
				city:       row[18],
				state:      row[19],
				postalCode: row[20],
				country:    row[21],
				phone:      row[22],
				note:       row[23],
			},
			order: Order{
				vol1SofCover:            row[2],
				extraVol2SofCover:       row[3],
				vol1ExtraBundle:         row[4],
				vol2SofCover:            row[5],
				doubleSidedBookmark:     row[6],
				vol2Booklet:             row[7],
				signedBookplate:         row[8],
				vol1SofCoverBook:        row[9],
				acrylicStandZevAviKeero: row[10],
				sketch:                  row[11],
				vol1ClockworkBook:       row[12],
				keeroKeychain:           row[13],
				vol1SignedBook:          row[14],
				acrylic:                 acrylics,
			},
		}

		orders = append(orders, newOrder)
	}

	// fmt.Print(orders)
	newSheet()
}

func acrylicSplit(a string) Acrylic {
	var acrylics Acrylic
	if a == "" {
		return acrylics
	}
	v := strings.Split(a, " | ")

	for _, e := range v {
		switch e {
		case "Big Keero":
			acrylics.bigKeero = 1
		case "Aline":
			acrylics.aline = 1
		case "Rhaine":
			acrylics.rhaine = 1
		case "Spixx":
			acrylics.spixx = 1
		case "Draven":
			acrylics.draven = 1
		case "Vespera":
			acrylics.vespera = 1
		case "Gabe":
			acrylics.gabe = 1
		case "Sash":
			acrylics.sash = 1
		case "Isaac":
			acrylics.isaac = 1
		case "Felice":
			acrylics.felice = 1
		case "The Dark Queen":
			acrylics.theDarkQueen = 1
		default:
			acrylics.other = 1
		}
	}

	return acrylics
}

func newSheet() {
	f := excelize.NewFile()
	index := f.NewSheet("Orders")
	row := 1

	for _, e := range orders {
		v := reflect.ValueOf(e.order)
		for i := 0; i < v.NumField(); i++ {

			if v.Type().Field(i).Name == "acrylic" {
				av := reflect.ValueOf(e.order.acrylic)
				for in := 0; in < av.NumField(); in++ {

					if av.Field(in).Int() != 0 {
						cell := "A" + fmt.Sprint(row)
						sku := "acrylic-" + av.Type().Field(in).Name

						err := f.SetSheetRow("Orders", cell, &[]interface{}{e.backerUID, sku, "1", " ", e.fullName, " ", e.address.address1, e.address.address2, e.address.city, e.address.state, e.address.postalCode, e.address.phone, e.address.country, e.email, " ", e.address.note})
						if err != nil {
							log.Fatal(err)
						}
						row++
					}
				}

			} else if v.Field(i).String() != "0" {
				cell := "A" + fmt.Sprint(row)

				err := f.SetSheetRow("Orders", cell, &[]interface{}{e.backerUID, v.Type().Field(i).Name, "1", " ", e.fullName, " ", e.address.address1, e.address.address2, e.address.city, e.address.state, e.address.postalCode, e.address.phone, e.address.country, e.email, " ", e.address.note})
				if err != nil {
					log.Fatal(err)
				}

				row++
			}
		}

		row++
	}

	err := f.InsertRow("Orders", 1)
	if err != nil {
		log.Fatal(err)
	}

	// OrderNumber	/ SKU	/ Quantity /	ShippingMethod /	FullName /	Company /	Address1 /	Address2 /	City /	State /	Zip	/ Phone /	Country	/ Email /	GiftMessage /	address Note / ParcelInsurance /
	err = f.SetSheetRow("Orders", "A1", &[]interface{}{"OrderNumber", "SKU", "Quantity", "ShippingMethod", "FullName", "Company", "Address1", "Address2", "City", "State", "Zip", "Phone", "Country", "Email", "GiftMessage", "Shipping Note", "ParcelInsurance"})
	if err != nil {
		log.Fatal(err)
	}

	f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	if err := f.SaveAs("Fufillment.xlsx"); err != nil {
		log.Fatal(err)
	}
}
