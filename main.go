package main

import (
	_ "embed"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"syscall"
)

//go:embed report_template.xlsm
var reportTemplate []byte

//go:embed rules.csv
var rules []byte

//go:embed ACSC_Hardening_Check.ps1
var script []byte

func main() {
	if runtime.GOOS == "windows" && !isAdminWindows() {
		fmt.Println("Requesting administrator privileges...")
		cmd := exec.Command("powershell", "Start-Process", "-Verb", "runas", os.Args[0])
		err := cmd.Run()
		if err != nil {
			fmt.Println("Error requesting administrator privileges:", err)
		}
		os.Exit(1)
	}

	execDir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		log.Fatal(err)
	}

	reportTemplatePath := filepath.Join(execDir, "report_template.xlsm")
	rulesPath := filepath.Join(execDir, "rules.csv")
	scriptPath := filepath.Join(execDir, "ACSC_Hardening_Check.ps1")

	// Write files to disk
	err = ioutil.WriteFile(reportTemplatePath, reportTemplate, 0644)
	if err != nil {
		log.Fatal(err)
	}

	err = ioutil.WriteFile(rulesPath, rules, 0644)
	if err != nil {
		log.Fatal(err)
	}

	err = ioutil.WriteFile(scriptPath, script, 0644)
	if err != nil {
		log.Fatal(err)
	}

	// Execute PowerShell script with admin privileges
	cmd := exec.Command("powershell", "-ExecutionPolicy", "Bypass", "-File", scriptPath)
	cmd.Stdin = os.Stdin
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	err = cmd.Run()
	if err != nil {
		log.Fatal(err)
	}

	// Delete the PowerShell script
	err = os.Remove(scriptPath)
	if err != nil {
		log.Fatal(err)
	}

}

func isAdminWindows() bool {
	mod := syscall.MustLoadDLL("shell32.dll")
	proc := mod.MustFindProc("IsUserAnAdmin")
	ret, _, _ := proc.Call()
	return ret != 0
}
