using System.Management;

public static class HardwareInfo
{
    /// Retrieving Motherboard Manufacturer.
    public static string GetBoardMaker()
    {
        var searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
        foreach (ManagementObject wmi in searcher.Get())
        {
            try
            {
                return wmi.GetPropertyValue("Manufacturer").ToString();
            }
            catch { }
        }
        return "Board Maker: Unknown";
    }
    /// Retrieving Motherboard Product Id.
    public static string GetBoardProductId()
    {
        var searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
        foreach (ManagementObject wmi in searcher.Get())
        {
            try
            {
                return wmi.GetPropertyValue("Product").ToString();
            }
            catch { }
        }
        return "Product: Unknown";
    }
    /// Retrieving BIOS Serial No.
    public static string GetBIOSserNo()
    {
        var searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BIOS");
        foreach (ManagementObject wmi in searcher.Get())
        {
            try
            {
                return wmi.GetPropertyValue("SerialNumber").ToString();
            }
            catch { }
        }
        return "BIOS Serial Number: Unknown";
    }
    /// Retrieving OS Info
    public static string GetOSInformation()
    {
        var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
        foreach (ManagementObject wmi in searcher.Get())
        {
            try
            {
                return ((string)wmi["Caption"]).Trim();
            }
            catch { }
        }
        return "BIOS Maker: Unknown";
    }
}