﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace ejemploColumpios1
{
    class Program
    {
        static Semaphore swings = new Semaphore(6, 6); // 6 columpios

        // Lista de nombres de los niños
        static string[] kids = {
        "Ana", "Ben", "Carlos", "Diana", "Eduardo",
        "Fernanda", "Gabriel", "Hugo", "Irene", "Javier",
        "Karen", "Luis", "Marta", "Nicolas", "Olivia",
        "Pablo", "Quintin", "Rosa", "Sofia", "Tomas"
    };
        static void Main(string[] args)
        {
            Console.WriteLine("Simulación de niños usando columpios con Thread Pool...");

            // Colocar cada niño en la cola de tareas del Thread Pool
            for (int i = 0; i < kids.Length; i++)
            {
                ThreadPool.QueueUserWorkItem(new WaitCallback(UseSwing), kids[i]);
            }

            // Esperar a que el usuario presione una tecla para finalizar el programa
            Console.ReadLine();
        }

        // Método que representa el uso del columpio por parte de un niño
        static void UseSwing(object kid)
        {
            string kidName = (string)kid;

            Console.WriteLine($"{kidName} está esperando para usar un columpio.");

            // Intentar entrar en el columpio (adquirir un columpio)
            swings.WaitOne();

            Console.WriteLine($"{kidName} está usando un columpio en hilo {Thread.CurrentThread.ManagedThreadId}.");

            // Simular tiempo de uso del columpio
            Thread.Sleep(new Random().Next(1000)); // Simular entre 2 y 5 segundos de uso

            Console.WriteLine($"{kidName} ha terminado de usar el columpio y lo ha dejado libre.");

            // Liberar el columpio
            swings.Release();
        }
    }
}
