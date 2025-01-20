#Requires AutoHotkey v2.0

Class PrinterManager {
    static ServerInstance := ''
    __New() {
        if PrinterManager.ServerInstance {
            return PrinterManager.ServerInstance
        }
        PrinterManager.ServerInstance := ComObject('WbemScripting.SWbemLocator').ConnectServer('.', 'root\cimv2')
    }
    GetPrinters() {
        printers := []
        try {
            query := PrinterManager.ServerInstance.ExecQuery('SELECT Name FROM Win32_Printer')

            for printer in query {
                try {
                    if printer.Name {
                        printers.Push(printer.Name)
                    }
                }
            }
        }
        return printers
    }
    GetDefaultPrinter() {
        try {
            query := PrinterManager.ServerInstance.ExecQuery('SELECT Name FROM Win32_Printer WHERE Default = True')
            for printer in query {
                return printer.Name
            }
        }
        return ''
    }
    GetPrinterMap() {
        printers := []
        try {
            query := PrinterManager.ServerInstance.ExecQuery('SELECT Name, Default FROM Win32_Printer')
            for printer in query {
                try {
                    if printer.Name {
                        printers.Push(Map('Name', printer.Name, 'Default', printer.Default))
                    }
                }
            }
        }
        return printers
    }
    SetDefaultPrinter(printerName) {
        if DllCall('Winspool.drv\SetDefaultPrinterW', 'Str', printerName) {
            return true
        }
        return false
    }
    PrinterExists(printerName) {
        if DllCall('Winspool.drv\OpenPrinterW', 'Str', printerName, 'Ptr*', &hPrinter := 0, 'Ptr', 0) {
            DllCall('Winspool.drv\ClosePrinter', 'Ptr', hPrinter)
            return true
        }
        return false
    }
    PrintTestPage(printerName) {
        if !DllCall('Winspool.drv\OpenPrinterW', 'Str', printerName, 'Ptr*', &hPrinter := 0, 'Ptr', 0) {
            return false
        }
        result := DllCall('Winspool.drv\PrinterProperties', 'Ptr', 0, 'Ptr', hPrinter)
        DllCall('Winspool.drv\ClosePrinter', 'Ptr', hPrinter)
        return result
    }
}