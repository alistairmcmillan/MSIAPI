// Quick and dirty demonstration of using MSI API to extract installed product info 
// Alistair McMillan
// 2014 July 4
//
// HOWTO COMPILE
// Download latest Mingw-w32 toolchain from http://www.drangon.org/mingw/
// Extract to c:\MinGW
// Add c:\MinGW\bin to PATH environment variable
// Compile with gcc msiapi.c -o msiapi.exe -l msi
//
// HOWTO RUN
// Run from Command Prompt with admin rights
// or
// Double-click in Windows Explorer if you run with admin rights
// or
// Right-click on msiapi.exe and choose "RunAs" or "Run as administrator" to be prompted for admin rights

#include <stdio.h>
#include <windows.h>
#include <conio.h>
#include <msi.h>

#define _UNICODE

CONSOLE_SCREEN_BUFFER_INFO csbi;
HANDLE hStdOutput;
BOOL bUsePause;

LPTSTR MsiQueryProperty(LPCTSTR szProductCode, 
                        LPCTSTR szUserSid,
                        MSIINSTALLCONTEXT dwContext,
                        LPCTSTR szProperty)
{
    LPTSTR lpValue[1024] = {0};

    DWORD pcchValue = 0;
    UINT retValue = MsiGetProductInfoEx(szProductCode, szUserSid, dwContext, szProperty, NULL, &pcchValue);

    if(retValue == ERROR_SUCCESS)
    {
        pcchValue++;
        retValue = MsiGetProductInfoEx( szProductCode, szUserSid, dwContext, szProperty, lpValue, &pcchValue);
    }

    return lpValue;
}

int __cdecl main(int argc, const char* argv[])
{
    UINT retValue = 0;
    DWORD dwIndex = 0;
    TCHAR szInstalledProductCode[39] = {0};
    TCHAR szSid[128] = {0};
    DWORD cchSid;
    MSIINSTALLCONTEXT dwInstalledContext;

    // Check whether we're being run from a Command Prompt
    // or from elsewhere. We want to pause if it's the latter.
    hStdOutput = GetStdHandle(STD_OUTPUT_HANDLE);
    if (!GetConsoleScreenBufferInfo(hStdOutput, &csbi))
    {
        printf("GetConsoleScreenBufferInfo failed: %d\n", GetLastError());
        return;
    }
    bUsePause = ((!csbi.dwCursorPosition.X) && (!csbi.dwCursorPosition.Y));

    // Enumerate list of products
    // Loop until call to MsiEnumProductsEx fails
    do {
        cchSid = sizeof(szSid)/sizeof(szSid[0]);

        // Call MsiEnumProductsEx on index
        retValue = MsiEnumProductsEx(NULL,
                                    "s-1-1-0",
                                    MSIINSTALLCONTEXT_USERMANAGED | MSIINSTALLCONTEXT_USERUNMANAGED | MSIINSTALLCONTEXT_MACHINE,
                                    dwIndex,
                                    szInstalledProductCode,
                                    &dwInstalledContext,
                                    szSid,
                                    &cchSid);

        // If call to MsiEnumProductsEx succeeds then query for information on that product
        // And write info to screen
        if(retValue == ERROR_SUCCESS)
        {
            printf("%s\n", MsiQueryProperty( szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_INSTALLEDPRODUCTNAME));
            printf(" - HelpLink: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_HELPLINK));
            printf(" - HelpTelephone: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_HELPTELEPHONE));
            printf(" - InstallDate: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_INSTALLDATE));
        //    printf(" - InstalledLanguage: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_INSTALLEDLANGUAGE)); // Not supported in Installer 4.5 and earlier
            printf(" - InstalledProductName: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_INSTALLEDPRODUCTNAME));
            printf(" - InstallLocation: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_INSTALLLOCATION));
            printf(" - InstallSource: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_INSTALLSOURCE));
            printf(" - LocalPackage: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_LOCALPACKAGE));
            printf(" - Publisher: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_PUBLISHER));
            printf(" - UrlInfoAbout: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_URLINFOABOUT));
            printf(" - UrlUpdateInfo: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_URLUPDATEINFO));
            printf(" - VersionMinor: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_VERSIONMINOR));
            printf(" - VersionMajor: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_VERSIONMAJOR));
            printf(" - VersionString: %s\n", MsiQueryProperty(szInstalledProductCode, cchSid == 0 ? NULL : szSid, dwInstalledContext, INSTALLPROPERTY_VERSIONSTRING));
            printf("\n");

            dwIndex++;
        }

    } while(retValue == ERROR_SUCCESS);

    printf("Found %d packages.\n", dwIndex);

    // Convert return code from MsiEnumProductsEx to actual error message
    LPVOID lpMsgBuf;
    FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                NULL,
                retValue,
                MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                (LPTSTR) &lpMsgBuf,
                0, NULL );

    // Write error message to screen
    printf("Ended with return code %d: %s", retValue, lpMsgBuf);

    // If return code was access denied then advice to run with admin rights
    if (retValue == ERROR_ACCESS_DENIED) {
        printf("\nQuerying MSI requires admin rights.\nPlease try again using an account with admin rights.\n");
    }

    // Pause if we weren't launched from a Command Prompt
    if (bUsePause)
    {
        int ch;
        printf("\nPress any key to exit...\n");
        ch = getch();
    }

    return 0;
}
