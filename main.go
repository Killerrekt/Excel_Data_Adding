package main

import (
	"context"
	"flag"
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
	"go.mongodb.org/mongo-driver/mongo"
	"go.mongodb.org/mongo-driver/mongo/options"
)

type Data struct {
	ReferenceNo          string `bson:"reference_no"`
	Name                 string `bson:"name"`
	Email                string `bson:"email"`
	MobileNo             string `bson:"mobile_no"`
	Address              string `bson:"address"`
	Country              string `bson:"country"`
	PaperID              string `bson:"paper_id"`
	IEEEMembershipID     string `bson:"ieee_membership_id"`
	ParticipantCategory  string `bson:"participant_category"`
	RegistrationCategory string `bson:"registration_category"`
	RegistrationDate     string `bson:"registration_date"`
	TransactionID        string `bson:"transaction_id"`
	InvoiceNo            string `bson:"invoice_no"`
	AmountPaid           string `bson:"amount_paid"`
	Event                string `bson:"event"`
}

var Mongo, DB *string

func main() {

	FileName := flag.String("f", "", "")
	Mongo = flag.String("m", "mongodb://localhost:27017/", "")
	DB = flag.String("db", "", "")
	Event := flag.String("e", "", "")

	flag.Parse()
	if *FileName == "" {
		log.Fatalln("Failed to mention the filename")
	}

	if *DB == "" {
		log.Fatalln("Didn't mention the DB")
	}

	if *Event == "" {
		log.Fatalln("Didn't pass any event")
	}

	file, err := excelize.OpenFile(*FileName + ".xlsx")
	if err != nil {
		log.Println("Failed to open the file")
	}
	defer func() {
		if err := file.Close(); err != nil {
			log.Println(err)
		}
	}()
	rows, err := file.GetRows("Worksheet")

	ConnectToMongo()

	for _, v := range rows[1:] {
		var temp Data
		temp.ReferenceNo = v[0]
		temp.Name = v[1]
		temp.Email = v[2]
		temp.MobileNo = v[3]
		temp.Address = v[4]
		temp.Country = v[5]
		temp.PaperID = v[6]
		temp.IEEEMembershipID = v[7]
		temp.ParticipantCategory = v[8]
		temp.RegistrationCategory = v[9]
		temp.RegistrationDate = v[10]
		temp.TransactionID = v[11]
		temp.InvoiceNo = v[12]
		temp.AmountPaid = v[13]
		temp.Event = *Event
		Col.InsertOne(context.TODO(), temp)
	}

	fmt.Println("ALL Data has been entered")
}

var Col *mongo.Collection

func ConnectToMongo() {
	var err error
	mongodb, err := mongo.Connect(context.TODO(), options.Client().ApplyURI(*Mongo))
	if err != nil {
		log.Fatalln("Failed to connect to the mongo : " + err.Error())
	}
	DB := mongodb.Database(*DB)
	Col = DB.Collection("Excel_Data")
	log.Println("Successfully connected to the MongoDB")
}
