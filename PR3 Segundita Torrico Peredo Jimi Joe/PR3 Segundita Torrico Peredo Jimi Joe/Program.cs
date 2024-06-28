using System;
using System.Threading;

class Program
{
    static int sum1, sum2, sum3, sum4;
    static AutoResetEvent resetEvent1 = new AutoResetEvent(false);
    static AutoResetEvent resetEvent2 = new AutoResetEvent(false);
    static AutoResetEvent resetEvent3 = new AutoResetEvent(false);
    static AutoResetEvent resetEvent4 = new AutoResetEvent(false);

    static void Main(string[] args)
    {
        Console.Write("Ingrese el inicio del rango: ");
        int ini = int.Parse(Console.ReadLine());

        Console.Write("Ingrese el fin del rango: ");
        int fin = int.Parse(Console.ReadLine());

        if (ini > fin)
        {
            Console.WriteLine("El inicio del rango debe ser menor o igual al fin del rango.");
            return;
        }

        int interv = (fin - ini + 1) / 4;

        ThreadPool.QueueUserWorkItem(state => Suminterv(ini, ini + interv - 1, ref sum1, resetEvent1));
        ThreadPool.QueueUserWorkItem(state => Suminterv(ini + interv, ini + 2 * interv - 1, ref sum2, resetEvent2));
        ThreadPool.QueueUserWorkItem(state => Suminterv(ini + 2 * interv, ini + 3 * interv - 1, ref sum3, resetEvent3));
        ThreadPool.QueueUserWorkItem(state => Suminterv(ini + 3 * interv, fin, ref sum4, resetEvent4));

        WaitHandle.WaitAll(new WaitHandle[] { resetEvent1, resetEvent2, resetEvent3, resetEvent4 });

        int totalSum = sum1 + sum2 + sum3 + sum4;

        Console.WriteLine($"La sumatoria del rango {ini} a {fin} es: {totalSum}");
        Console.ReadLine();
    }

    static void Suminterv(int ini, int fin, ref int sum, AutoResetEvent resetEvent)
    {
        sum = 0;
        for (int i = ini; i <= fin; i++)
        {
            sum += i;
        }
        resetEvent.Set();
    }
}

