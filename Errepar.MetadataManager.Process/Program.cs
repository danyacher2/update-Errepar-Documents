try
{
    Console.WriteLine("Hello, World!");
    // Aquí tu lógica...
}
catch (Exception ex)
{
    Console.Error.WriteLine("Excepción no controlada: " + ex);
    Environment.ExitCode = 1;
}
finally
{
    Console.WriteLine("Proceso finalizado con código " + Environment.ExitCode + ". Presiona una tecla para cerrar...");
    Console.ReadKey();
}