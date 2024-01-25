package main

import (
	"context"
	"flag"
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
	"go.mongodb.org/mongo-driver/bson"
	"go.mongodb.org/mongo-driver/mongo"
	"go.mongodb.org/mongo-driver/mongo/options"
)

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
		temp := bson.M{}
		count := 0
		for _, val := range v {
			temp[rows[0][count]] = val
			count++
		}
		temp["Event"] = *Event
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
