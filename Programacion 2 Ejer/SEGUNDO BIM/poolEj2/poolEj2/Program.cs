using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace poolEj2
{
    class Program
    {
        // Lista de materias
        static string[] subjects = {
        "Matemáticas", "Física", "Química", "Biología",
        "Historia", "Geografía", "Lengua", "Inglés",
        "Arte", "Informática","Matemáticas2", "Física2", "Química2", "Biología2",
        "Historia2", "Geografía2", "Lengua2", "Inglés2",
        "Arte2", "Informática2"
    };
        static void Main(string[] args)
        {
            int processorCount = Environment.ProcessorCount;
            Console.WriteLine("Número de hilos disponibles en el CPU: " + processorCount);
            Console.WriteLine("Iniciando estudio de materias utilizando Thread Pool...");

            // Colocar cada materia en la cola de tareas del Thread Pool
            for (int i = 0; i < subjects.Length; i++)
            {
                ThreadPool.QueueUserWorkItem(new WaitCallback(StudySubject), subjects[i]);
            }

            // Esperar a que el usuario presione una tecla para finalizar el programa
            Console.ReadLine();
        }

        // Método que representa la tarea de estudiar una materia
        static void StudySubject(object subject)
        {
            string subjectName = (string)subject;
            Console.WriteLine($"Estudiando {subjectName} en hilo {Thread.CurrentThread.ManagedThreadId}...");

            // Simular tiempo de estudio con un retardo
            Thread.Sleep(new Random().Next(1000, 5000)); // Simular entre 1 y 5 segundos de estudio

            Console.WriteLine($"Completado el estudio de {subjectName} en hilo {Thread.CurrentThread.ManagedThreadId}.");
        }
    }
}
