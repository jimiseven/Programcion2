using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2

    //using System;

// Clase base genérica
public class Animal<T>
{
    public T Nombre { get; set; }

    public Animal(T nombre)
    {
        Nombre = nombre;
    }

    public virtual void EmitirSonido()
    {
        Console.WriteLine("Haciendo ruido...");
    }
}

// Clase derivada que hereda de Animal<T>
public class Perro<T> : Animal<T>
{
    public Perro(T nombre) : base(nombre)
    {
    }

    // Método específico de la clase derivada
    public void Ladrar()
    {
        Console.WriteLine($"{Nombre} está ladrando: ¡Guau Guau!");
    }
}

class Program
{
    static void Main(string[] args)
    {
        // Crear un perro con nombre de tipo string
        var miPerro = new Perro<string>("Fido");

        // Acceder a propiedades y métodos de la clase base y derivada
        Console.WriteLine($"Nombre del perro: {miPerro.Nombre}");
        miPerro.EmitirSonido(); // Método de la clase base
        miPerro.Ladrar(); // Método de la clase derivada
    }
}