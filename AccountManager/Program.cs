using OfficeOpenXml;

namespace AccountManager
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            CheckVersionValidity();

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new AccountManager());
        }
        private static void CheckVersionValidity()
        {
            DateTime expirationDate = new DateTime(2025, 9, 1);
            if (DateTime.Now >= expirationDate)
            {
                MessageBox.Show("Текущая версия программы больше не действительна. Обратитесь к разработчику для обновления.",
                                "Версия не действительна",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                Environment.Exit(0);
            }
        }
    }
}