namespace NpoiExample
{
    class Program
    {
        static void Main()
        {
            var path = @"./consultaMovimientosEfectivo(04-feb-16).xls";
            var reader = new ExcelReader();
            reader.Read(path);
        }
    }
}
