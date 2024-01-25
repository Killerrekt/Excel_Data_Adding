package main

import (
	"context"
	"fmt"
	"log"

	"github.com/spf13/viper"
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
}

func main() {
	viper.SetConfigFile(".env")
	viper.AutomaticEnv()
	viper.ReadInConfig()
	file, err := excelize.OpenFile("registration_report_25_01_2024.xlsx") //vchange the name based on the excel sheet name
	if err != nil {
		log.Println("Failed to open the file")
	}
	defer func() {
		if err := file.Close(); err != nil {
			log.Println(err)
		}
	}()
	rows, err := file.GetRows("Worksheet")
	//fmt.Println(rows[1:])

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
		Col.InsertOne(context.TODO(), temp)
	}

	fmt.Println("ALL Data has been entered")
}

var Col *mongo.Collection

func ConnectToMongo() {
	var err error
	mongodb, err := mongo.Connect(context.TODO(), options.Client().ApplyURI(viper.GetString("MONGO_URI")))
	if err != nil {
		log.Fatalln("Failed to connect to the mongo : " + err.Error())
	}
	DB := mongodb.Database(viper.GetString("MONGO_DATABASE"))
	Col = DB.Collection("Excel_Data")
	log.Println("Successfully connected to the MongoDB")
}
