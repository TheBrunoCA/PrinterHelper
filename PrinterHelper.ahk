#Requires AutoHotkey v2.0
/*
 * @description PrinterHelper is a class designed to simplify printer management.
 * @author TheBrunoCA
 */

Class PrinterHelper {

    /**
     * @description Sets the logging callback
     * @param {function} callback - The callback to be called when logging, must be a function with a single string parameter
     */
    static LoggingCallback := ''

    /**
     * @description Retrieves all available printers using PowerShell.
     * @return Array of printer names.
     */
    static GetPrinters() {
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call("Getting printers")
        printers := []
        try {
            wmi := ComObject("WbemScripting.SWbemLocator").ConnectServer(".", "root\cimv2")
            query := wmi.ExecQuery("SELECT Name FROM Win32_Printer")
            
            for printer in query {
                try {
                    if printer.Name {
                        printers.Push(printer.Name)
                        if PrinterHelper.LoggingCallback
                            PrinterHelper.LoggingCallback.Call("Got printer: " printer.Name)
                    }
                }
            }
        }
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call('Got all printers')
        return printers
    }

    /**
     * @description Retrieves the default printer using PowerShell.
     * @return Name of the default printer.
     */
    static GetDefaultPrinter() {
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call("Getting default printer")
        try {
            wmi := ComObject("WbemScripting.SWbemLocator").ConnectServer(".", "root\cimv2")
            query := wmi.ExecQuery("SELECT Name FROM Win32_Printer WHERE Default = True")
            
            for printer in query {
                if PrinterHelper.LoggingCallback
                    PrinterHelper.LoggingCallback.Call("Got default printer: " printer.Name)
                return printer.Name
            }
        }
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call('Cound not get default printer')
        return ""
    }

    /**
     * @description Retrieves a map of all printers with their default status using PowerShell.
     * @return Array of Maps containing printer details.
     */
    static GetPrinterMap() {
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call("Getting printer map, format is Name - Default")
        printers := []
        try {
            wmi := ComObject("WbemScripting.SWbemLocator").ConnectServer(".", "root\cimv2")
            query := wmi.ExecQuery("SELECT Name, Default FROM Win32_Printer")
            
            for printer in query {
                try {
                    if printer.Name {
                        printers.Push(Map(
                            "Name", printer.Name,
                            "Default", printer.Default
                        ))
                        if PrinterHelper.LoggingCallback
                            PrinterHelper.LoggingCallback.Call("Got printer map: " printer.Name " - " printer.Default)
                    }
                }
            }
        }
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call('Got all printers mapped')
        return printers
    }

    /**
     * @description Sets the default printer by name using PowerShell.
     * @param name (String) - The name of the printer to set as default.
     * @return Boolean - True if successful, false otherwise.
     */
    static SetDefaultPrinter(name) {
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call("Setting default printer: " name)
        if DllCall("Winspool.drv\SetDefaultPrinterW", "Str", name) {
            if PrinterHelper.LoggingCallback
                PrinterHelper.LoggingCallback.Call("Successfully set default printer: " name)
            return true
        }
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call("Failed to set default printer: " name)
        return false
    }

    /**
     * @description Checks if a specific printer exists.
     * @param name (String) - The name of the printer to check.
     * @return Boolean - True if the printer exists, false otherwise.
     */
    static PrinterExists(name) {
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call("Checking if printer exists: " name)
        if DllCall("Winspool.drv\OpenPrinterW", "Str", name, "Ptr*", &hPrinter := 0, "Ptr", 0) {
            DllCall("Winspool.drv\ClosePrinter", "Ptr", hPrinter)
            if PrinterHelper.LoggingCallback
                PrinterHelper.LoggingCallback.Call("Printer exists: " name)
            return true
        }
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call("Printer does not exist: " name)
        return false
    }

    /**
     * @description Prints a test page on a specific printer using PowerShell.
     * @param name (String) - The name of the printer.
     * @return Boolean - True if the test page was sent successfully.
     */
    static PrintTestPage(name) {
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call("Printing test page on printer: " name)
        if !DllCall("Winspool.drv\OpenPrinterW", "Str", name, "Ptr*", &hPrinter := 0, "Ptr", 0) {
            if PrinterHelper.LoggingCallback
                PrinterHelper.LoggingCallback.Call("Failed to open printer: " name)
            return false
        }

        result := DllCall("Winspool.drv\PrinterProperties", "Ptr", 0, "Ptr", hPrinter)
        
        DllCall("Winspool.drv\ClosePrinter", "Ptr", hPrinter)
        
        if PrinterHelper.LoggingCallback
            PrinterHelper.LoggingCallback.Call("Printed test page on printer: " name)
        return result
    }
}
