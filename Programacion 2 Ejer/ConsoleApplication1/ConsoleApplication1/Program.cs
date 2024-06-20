using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {

            //AlmcenObj<DateTime> info = new AlmcenObj<DateTime>(4);
            //info.agregar(new DateTime(2015, 12, 25));
            //info.agregar(new DateTime(2016, 12, 25));
            //info.agregar(new DateTime(2017, 12, 25));
            //info.agregar(new DateTime(2018, 12, 25));
            //DateTime infoFecha = info.getEle(2);
            //Console.WriteLine(infoFecha);


            //AlmcenObj<String> info = new AlmcenObj<String>(4);
            //info.agregar("jimi");
            //info.agregar("joe");
            //info.agregar("gino");
            //String nombreEmpleado = info.getEle(2);
            //Console.WriteLine(nombreEmpleado);


            AlmcenObj<Empleado> info = new AlmcenObj<Empleado>(4);
            info.agregar(new Empleado(1000));
            info.agregar(new Empleado(2000));
            info.agregar(new Empleado(6000));
            info.agregar(new Empleado(4000));
            Empleado salarioEmpleado = (Empleado)info.getEle(2);
            Console.WriteLine(salarioEmpleado.getSalario());
            
        }
    }
    //un generico se genera con una T mayuscula por convenciones
    class AlmcenObj <T> {
        public AlmcenObj(int z) {
            datosElemento = new T[z];
        }

        public void agregar(T obj) {
            datosElemento[i] = obj;
            i++;
        }

        public T getEle(int i) {
            return datosElemento[i];
        }

        private T[] datosElemento;

        private int i = 0;

    }


    class Empleado {
        private double salario;

        public Empleado(double salario) {
            this.salario = salario;
        }

        public double getSalario() {
            return salario;
        }
    }
}
