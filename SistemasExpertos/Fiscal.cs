using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SistemasExpertos
{
    class Fiscal
    {
        private string _id_fiscal;
        private string _apell;//usado hoja 1, apellido de fiscal
        private string _nomb;//usado hoja 1, nombre de fiscal
        private double _pend;
        private double _calif;
        private double _preli;
        private double _prepa;
        private double _inter;
        private double _juzga;
        private double _deriv;//derivados
        private double _archi;//archivo sin consentir
        private double _archicon;//archivo consentido
        private double _pnp;
        private double _prov;//reserva provicional
        private double _acum;//acumulado
        private double _rpnp;//reserva pnp
        private double _asig;//ingresos asignados
        private double _expe;//antiguo codigo
        private double _iexpe;//antiguo codigo
        private double _apel;
        private double _impug;
        private double _cons;
        private double _excl;
        private double _senten;//sentencias
        private double _quejas;
        private double _ipre;//antiguo codigo
        private double _sobreseim;//sobreseimiento sin consentir
        private double _sobreseimcon;//sobreseimiento consentido
        private double _acuerep;//acuerdo reparatorio
        private double _princoport;//principio de oportunidad
        private double _id_depen_mpub;
        private string _id_esp;
        private double _pdom;
        private double _qpdom;

        public string Id_fiscal { get => _id_fiscal; set => _id_fiscal = value; }
        public string Apell { get => _apell; set => _apell = value; }
        public string Nomb { get => _nomb; set => _nomb = value; }
        public double Pend { get => _pend; set => _pend = value; }
        public double Calif { get => _calif; set => _calif = value; }
        public double Preli { get => _preli; set => _preli = value; }
        public double Prepa { get => _prepa; set => _prepa = value; }
        public double Inter { get => _inter; set => _inter = value; }
        public double Juzga { get => _juzga; set => _juzga = value; }
        public double Deriv { get => _deriv; set => _deriv = value; }
        public double Archi { get => _archi; set => _archi = value; }
        public double Archicon { get => _archicon; set => _archicon = value; }
        public double Pnp { get => _pnp; set => _pnp = value; }
        public double Prov { get => _prov; set => _prov = value; }
        public double Acum { get => _acum; set => _acum = value; }
        public double Rpnp { get => _rpnp; set => _rpnp = value; }
        public double Asig { get => _asig; set => _asig = value; }
        public double Expe { get => _expe; set => _expe = value; }
        public double Iexpe { get => _iexpe; set => _iexpe = value; }
        public double Apel { get => _apel; set => _apel = value; }
        public double Impug { get => _impug; set => _impug = value; }
        public double Cons { get => _cons; set => _cons = value; }
        public double Excl { get => _excl; set => _excl = value; }
        public double Senten { get => _senten; set => _senten = value; }
        public double Quejas { get => _quejas; set => _quejas = value; }
        public double Ipre { get => _ipre; set => _ipre = value; }
        public double Sobreseim { get => _sobreseim; set => _sobreseim = value; }
        public double Sobreseimcon { get => _sobreseimcon; set => _sobreseimcon = value; }
        public double Acuerep { get => _acuerep; set => _acuerep = value; }
        public double Princoport { get => _princoport; set => _princoport = value; }
        public double Id_depen_mpub { get => _id_depen_mpub; set => _id_depen_mpub = value; }
        public string Id_esp { get => _id_esp; set => _id_esp = value; }
        public double Pdom { get => _pdom; set => _pdom = value; }
        public double Qpdom { get => _qpdom; set => _qpdom = value; }
    }
}
