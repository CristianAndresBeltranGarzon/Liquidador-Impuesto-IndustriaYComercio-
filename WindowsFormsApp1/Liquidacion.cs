using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class Liquidacion
    {
        private int cedula = 0;
        private string nombre = "";

        public Liquidacion(int c,string s) {
            cedula = c;
            nombre = s;
        }
        public override string ToString()
        {
            return string.Format("la cedula es {0} el nombre es {1}", cedula, nombre);
        }
    }
}
