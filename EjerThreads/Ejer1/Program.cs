using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;//llamamos labreria +


//metodos importantes
//join para sincronizar termina una tarea y despues empieza con otra
//block para bloquear

namespace Ejer1
{
    class Program
    {
        static void Main(string[] args)
        {
            Thread t = new Thread(MetodoSaludo);//utiilizamos el metodosaludo para thread
            t.Start();//iniciamos el hilo
            t.Join();// hasta que no termine el t no hace t2
            Thread t2 = new Thread(MetodoSaludo);//utiilizamos el metodosaludo para thread
            t2.Start();//iniciamos el hilo
            t2.Join();
            Console.WriteLine("hola");
            Thread.Sleep(5000);//duerme el hilo
            Console.WriteLine("hola");
            Console.WriteLine("hola");
        }
        static void MetodoSaludo() {
            Console.WriteLine("hola desde thread");
            Console.WriteLine("hola desde thread");
        }
    }
}
