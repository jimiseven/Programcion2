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
            AlmcenObj info = new AlmcenObj(4);

            info.agregar("jimi");
            info.agregar("joe");
            info.agregar("gino");

            String nombrePersona = (String)info.getEle(2);
            //realizando casting
            Console.WriteLine(nombrePersona);
        }
    }

    class AlmcenObj
    {
        public AlmcenObj(int z)
        {
            datosElemento = new Object[z];
        }

        public void agregar(Object obj)
        {
            datosElemento[i] = obj;
            i++;
        }

        public Object getEle(int i)
        {
            return datosElemento[i];
        }

        private Object[] datosElemento;

        private int i = 0;

    }

}
