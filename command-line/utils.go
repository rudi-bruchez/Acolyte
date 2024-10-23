package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/joho/godotenv"
)

func EnvFileName(toolName string) string {
	return filepath.Join(GetConfigDir(), toolName, `.env`)
}

func EnvFileExists(toolName string) bool {
	if _, err := os.Stat(EnvFileName(toolName)); os.IsNotExist(err) {
		return false
	}
	return true
}

func LoadEnvFile(toolName string) {
	err := godotenv.Load(EnvFileName(toolName))
	if err != nil {
		log.Fatal(fmt.Sprintf("Error loading %s.env file: ", toolName), err)
	}
}

func EnsureEnvFile(toolName string, vars []string) {
	envFile := EnvFileName(toolName)

	var f *os.File
	var err error
	if !EnvFileExists(toolName) {
		f, err = os.Create(envFile)
		if err != nil {
			log.Fatal(fmt.Sprintf("Error creating %s.env file: ", toolName), err)
		}
	} else {
		f, err = os.Open(envFile)
		if err != nil {
			log.Fatal(fmt.Sprintf("Error opening %s.env file: ", toolName), err)
		}
		LoadEnvFile(toolName)
	}
	defer f.Close()

	envvars := make(map[string]string)

	for _, v := range vars {
		// f.WriteString(v + "\n")
		value := os.Getenv(v)

		if value == "" {
			fmt.Printf("Please enter the value of %s: ", v)
			fmt.Scanln(&value)
			os.Setenv(v, value)
			envvars[v] = value
		}
	}
	godotenv.Write(envvars, envFile)

}

func GetConfigDir() string {
	var homedir string
	var err error
	homedir, err = os.UserHomeDir()
	if err != nil {
		log.Fatal("Error getting user home directory: ", err)
		// return nil
	}

	return filepath.Join(homedir, ".config/stitch/")
}

func IsFlagPassed(name string) bool {
	found := false
	flag.Visit(func(f *flag.Flag) {
		if f.Name == name {
			found = true
		}
	})
	return found
}

func ListFilesInDirectory(dir string, recursive bool, f func(string)) {
	files, err := os.ReadDir(dir)
	if err != nil {
		log.Fatal(err)
	}

	for _, file := range files {
		if file.IsDir() && recursive {
			ListFilesInDirectory(filepath.Join(dir, file.Name()), recursive, f)
		} else {
			fmt.Println(file.Name())
			f(file.Name())
		}
	}
}
