namespace vbamc.Vba
{
    [Flags]
    public enum ProjectProtection
    {
        None = 0,
        UserProtected = 1,
        HostProtected = 2,
        VBEProtected = 4,
    }
}
