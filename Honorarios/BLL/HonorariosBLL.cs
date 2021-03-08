using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Honorarios.BLL
{
    class HonorariosBLL
    {
        BaseDatosEntities1 context;

        public List<string> GetListRut()
        {
            context = new BaseDatosEntities1();
            return (from l in context.HonorariosDatos select l.Rut).Distinct().ToList();
        }

        public List<HonorariosDatos> GetBoletas(string rut)
        {
            context = new BaseDatosEntities1();
            return (from l in context.HonorariosDatos where rut == l.Rut select l).ToList();
        }
    }
}
