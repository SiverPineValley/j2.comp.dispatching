package config

import (
	"fmt"

	"github.com/BurntSushi/toml"
	"js.comp.dispatching/src/models"
)

// Config is frame of config.toml
type Config struct {
	Target  map[string][]string `toml:"target"`
	Company map[string][]string `toml:"company"`
	J2      ColumnTitle         `toml:"j2"`
	Cj      ColumnTitle         `toml:"cj"`
	Gansun  ColumnTitle         `toml:"gansun"`
	Direct  map[string]string   `toml:"direct"`
}

// ColumnTitle keep each j2, cj column name.
type ColumnTitle struct {
	No            string `toml:"no"`
	Date          string `toml:"date"`
	LicensePlate  string `toml:"licensePlate"`
	TrailerPlate  string `toml:"trailerPlate"`
	Source        string `toml:"source"`
	Destination   string `toml:"destination"`
	Route         string `toml:"route"`
	LayoverNum    string `toml:"layoverNum"`
	CarType       string `toml:"carType"`
	Reference     string `toml:"reference"`
	TargetCompany string `toml:"targetCompany"`
	StartIdx      int    `toml:"startIndex"`
}

// Gansun is gansun route - company filtering config
type Gansun struct {
	Route map[string][][]string `toml:"route"`
}

// getColumnName is parsing function about config.toml.
func getColumnName(config *Config) {
	_, err := toml.DecodeFile("config.toml", &config)
	if err != nil {
		fmt.Println(models.InvalidTomlErr)
		return
	}
	return
}

// InitConfig is config Initalization funciton.
func InitConfig(config *Config) {
	getColumnName(config)
	return
}

// InitGansun is gansun config Initalization funciton.
func InitGansun(gansun *Gansun) {
	_, err := toml.DecodeFile("gansun.toml", &gansun)
	if err != nil {
		fmt.Println(models.InvalidTomlErr)
		return
	}
	return
}
