# IT Support Tool GUI

A WPF application for IT support staff to quickly get system information from a user's machine, and perform some common IT support tasks.

## Features

This tool provides a graphical user interface for the following functionalities:

*   **Display System Information:**
    *   Computer Name
    *   Username
    *   OS Version
    *   SSID (WiFi network) or Ethernet connection status
    *   IP Address
    *   VPN Status (specifically for FortiClient)
    *   VPN IP Address
    *   Last system restart time.

*   **One-Click Actions:**
    *   Copy all system information to the clipboard.
    *   Open a ticket in ServiceNow.
    *   Open a password reset page.
    *   Open "IT Messages" on SharePoint, with different URLs based on the user's country (Sweden, Norway, Denmark, or Group).
    *   Launch LogMeIn Rescue Calling Card for remote help.
    *   Run a PowerShell script to fix Wi-Fi issues by removing the 'StarkGroup' Wi-Fi profile.

*   **Printer Management:**
    *   Search for network printers (Ricoh and Zebra) based on a Site ID.
    *   Add selected network printers to the user's system.

## Prerequisites

To build and run this project, you will need the [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).

## Building and Running

1.  **Clone the repository:**
    ```sh
    git clone <repository-url>
    cd <repository-directory>
    ```

2.  **Build the project:**
    ```sh
    dotnet build
    ```

3.  **Run the application:**
    ```sh
    dotnet run --project ITSupportToolGUI
    ```

## Localization

The application supports the following languages:

*   Danish (da-DK)
*   English (en-US)
*   Norwegian (nb-NO)
*   Swedish (sv-SE)

The application's language is automatically determined based on the system's region settings.

## Dependencies

The project uses the following NuGet packages:

*   `System.Management`
*   `System.Management.Automation`
