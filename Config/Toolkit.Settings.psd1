@{
    LogRoot = 'Logs'

    Integrations = @{
        Microsoft365RepoZip = 'https://github.com/mallockey/Install-Microsoft365/archive/refs/heads/main.zip'
        WindowsMediaSupportPage = 'https://support.microsoft.com/en-us/windows/create-installation-media-for-windows-99a58364-8c02-206f-aa6f-40c3b507420d'
        Windows10MediaToolUrl = 'https://go.microsoft.com/fwlink/?LinkId=691209'
    }

    UpdateTools = @{
        SupportedManufacturers = @('Dell', 'HP', 'Lenovo')
        DellCommandUpdatePage = 'https://www.dell.com/support/home/en-us/drivers/driversdetails?driverid=5cr1y&oscode=wt64a&productcode=command-update&src=o'
        HPClientManagementPage = 'https://www.hp.com/us-en/solutions/client-management-solutions.html'
        HPImageAssistantPage = 'https://support.hp.com/us-en/drivers/consumers/hp-image-assistant'
        LenovoSystemUpdatePage = 'https://support.lenovo.com/us/en/solutions/ht037099'
        LenovoSystemUpdateDocs = 'https://docs.lenovocdrt.com/guides/sus/'
        WindowsUpdateApiDocs = 'https://learn.microsoft.com/en-us/windows/win32/api/wuapi/nf-wuapi-iupdatesearcher-search'
    }
}
