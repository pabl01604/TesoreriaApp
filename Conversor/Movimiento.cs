using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Conversor
{
    public class Movimiento
    {
        public DateTime Fecha { get; set; }
        public int SucOrigen { get; set; }

        public string DescSucursal { get; set; }
        public int CodOperativo { get; set; }
        public int Referencia { get; set; }
        public string Concepto { get; set; }
        public double ImporteEnPesos { get; set; }

        public Movimiento()
        {
       
        }

        public Movimiento(DateTime fecha, int sucursal, int cod, int refer, string concep, double importe )
        {
            this.Fecha = fecha;
            SucOrigen = sucursal;

            switch (SucOrigen)
            {
                case 0:  DescSucursal = "Casa Central";
                        break;
                case 64: DescSucursal="Villa Cabrera";
                    break;
                case 66: DescSucursal = "R. DE SANTA FE";
                    break;
                case 556: DescSucursal = "Los Boulevares Córdoba Capital";
                    break;
                default:  DescSucursal = "Desconocido";
                    break;
            }

            CodOperativo = cod;
            Referencia= refer;
            Concepto= concep;
            ImporteEnPesos= importe;
        }

    }
}
