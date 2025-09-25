using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.IO;

namespace AppGetInfo
{
    public partial class FormAppGetInfo : Form
    {        
        public FormAppGetInfo()
        {
            InitializeComponent();
            txtAlmacenamiento.ScrollBars = ScrollBars.Both;
            txtAlmacenamiento.WordWrap = false;
        }

        private void btnObtenerInformacion_Click(object sender, EventArgs e)
        {
            txtMotherBoard.Text = ObtenerInfoMotherboard();
            txtTarjetaGrafica.Text = ObtenerInformacionGrafica();
            txtRAM.Text = ObtenerInformacionTipoRAM();
            txtCantidadRAM.Text = ObtenerInformacionCantidadRAM();
            txtCantSlots.Text = ObtenerInformacionCantidadSlots();
            txtAlmacenamiento.Text = ObtenerInformacionDisco();
            txtProcesador.Text = ObtenerInformacionProcesador();
            txtSO.Text = ObtenerInformacionSO();
        }

        public string ObtenerInfoMotherboard()
        {
            ManagementObjectSearcher mos = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
            string motherboard = "";
            foreach (ManagementObject mo in mos.Get())
            {
                try
                {
                    motherboard = mo.GetPropertyValue("SerialNumber").ToString()
                + " | " + mo.GetPropertyValue("Manufacturer").ToString()
                + " | " + mo.GetPropertyValue("Product").ToString();
                }
                catch
                { }
            }
            return motherboard;
        }

        public string ObtenerInformacionGrafica()
        {
            string grafica = "";
            ManagementObjectSearcher myVideoObject = new ManagementObjectSearcher("select * from Win32_VideoController");

            foreach (ManagementObject obj in myVideoObject.Get())
            {
                grafica = obj["DriverVersion"] + " | " + obj["Name"] + " | " + obj["VideoProcessor"];
            }
            return grafica;
        }

        public static string FormatBytes(long bytes)
        {
            string[] Suffix = { "B", "KB", "MB", "GB", "TB" };
            int i;
            double dblSByte = bytes;
            for (i = 0; i < Suffix.Length && bytes >= 1024; i++, bytes /= 1024)
            {
                dblSByte = bytes / 1024.0;
            }

            return String.Format("{0:0.##} {1}", dblSByte, Suffix[i]);
        }

        public string ObtenerInformacionTipoRAM()
        {
            int type = 0;
            int smbiosType = 0;
            bool found = false;
            // Usar la ruta correcta para ManagementScope
            ManagementScope scope = new ManagementScope("root\\CIMV2");
            scope.Connect();
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_PhysicalMemory");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            foreach (ManagementObject queryObj in searcher.Get())
            {
                if (queryObj["SMBIOSMemoryType"] != null)
                {
                    smbiosType = Convert.ToInt32(queryObj["SMBIOSMemoryType"]);
                    found = true;
                    break;
                }
                else if (queryObj["MemoryType"] != null)
                {
                    type = Convert.ToInt32(queryObj["MemoryType"]);
                    found = true;
                    break;
                }
            }
            if (found)
            {
                if (smbiosType != 0)
                    return TypeString(smbiosType);
                else
                    return TypeString(type);
            }
            else
            {
                return "No disponible";
            }
        }

        private static string TypeString(int type)
        {
            string outValue = string.Empty;

            switch (type)
            {
                case 0x0: outValue = "Unknown"; break;
                case 0x1: outValue = "Other"; break;
                case 0x2: outValue = "DRAM"; break;
                case 0x3: outValue = "Synchronous DRAM"; break;
                case 0x4: outValue = "Cache DRAM"; break;
                case 0x5: outValue = "EDO"; break;
                case 0x6: outValue = "EDRAM"; break;
                case 0x7: outValue = "VRAM"; break;
                case 0x8: outValue = "SRAM"; break;
                case 0x9: outValue = "RAM"; break;
                case 0xa: outValue = "ROM"; break;
                case 0xb: outValue = "Flash"; break;
                case 0xc: outValue = "EEPROM"; break;
                case 0xd: outValue = "FEPROM"; break;
                case 0xe: outValue = "EPROM"; break;
                case 0xf: outValue = "CDRAM"; break;
                case 0x10: outValue = "3DRAM"; break;
                case 0x11: outValue = "SDRAM"; break;
                case 0x12: outValue = "SGRAM"; break;
                case 0x13: outValue = "RDRAM"; break;
                case 0x14: outValue = "DDR"; break;
                case 0x15: outValue = "DDR2"; break;
                case 0x16: outValue = "DDR2 FB-DIMM"; break;
                case 0x17: outValue = "Undefined 23"; break;
                case 0x18: outValue = "DDR3"; break;
                case 0x19: outValue = "FBD2"; break;
                case 0x1a: outValue = "DDR4"; break;
                default: outValue = "Undefined"; break;
            }

            return outValue;
        }
        public string ObtenerInformacionCantidadRAM()
        {
            string cantidadRAM = "";
            ObjectQuery wql = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(wql);
            ManagementObjectCollection results = searcher.Get();

            foreach (ManagementObject result in results)
            {
                long cantidadMB = long.Parse(result["TotalVisibleMemorySize"].ToString());
                string cantidadGB = FormatBytes(cantidadMB * 1024);
                cantidadRAM = cantidadGB + " (Disponible)";            
            }
            return cantidadRAM;
        }

        public string ObtenerInformacionCantidadSlots()
        {
            ManagementObjectSearcher searcher;
            string cantidadSlots = "";
            try
            {
                searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PhysicalMemoryArray");
                foreach (ManagementObject queryObj in searcher.Get())
                {
                    cantidadSlots = queryObj["MemoryDevices"].ToString();
                }
            }
            catch (ManagementException e)
            {
                System.IO.File.AppendAllText(@"ErrorAppAD.txt", "Opps... Tuvimos el siguiente error en consultaRAM: " + e.ToString());
            }            
            return cantidadSlots;
        }

        public string ObtenerInformacionDisco()
        {
            string disco = "";
            DriveInfo[] allDrives = DriveInfo.GetDrives();

            foreach (DriveInfo d in allDrives)
            {
                if (d.IsReady == true)
                {
                    disco += $"Disco: {d.Name} | Formato: {d.DriveFormat} | Espacio Disponible: {FormatBytes(d.TotalFreeSpace)} | Espacio Utilizado: {FormatBytes(d.TotalSize - d.TotalFreeSpace)} | Espacio Total: {FormatBytes(d.TotalSize)}\r\n";
                }
            }
            return disco;
        }

        public string ObtenerInformacionProcesador()
        {
            string procesador = "";
            ManagementObjectSearcher myProcessorObject = new ManagementObjectSearcher("select * from Win32_Processor");

            foreach (ManagementObject obj in myProcessorObject.Get())
            {
                procesador += obj["Name"];
            }
            return procesador;
        }

        public string ObtenerInformacionSO()
        {
            ManagementObjectSearcher myOperativeSystemObject = new ManagementObjectSearcher("select * from Win32_OperatingSystem");

            string so = "";
            foreach (ManagementObject obj in myOperativeSystemObject.Get())
            {
                so +=  obj["Caption"];
               
            }
            return so;
        }

        public string ObtenerInformacionTipoDisco(DriveInfo d)
        {
            try
            {
                var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_DiskDrive WHERE Index = {d.Name.Replace(":\\","")}");
                foreach (ManagementObject obj in searcher.Get())
                {
                    var mediaType = obj["MediaType"]?.ToString();
                    if (mediaType != null && mediaType.ToLower().Contains("ssd"))
                        return "SSD";
                    else if (mediaType != null && mediaType.ToLower().Contains("hdd"))
                        return "HDD";
                }
            }
            catch { }
            return "Desconocido";
        }

        public string ObtenerInformacionRed()
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True");
                foreach (ManagementObject obj in searcher.Get())
                {
                    sb.AppendLine($"Adaptador: {obj["Description"]} | IP: {((string[])(obj["IPAddress"]))[0]}");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Error obteniendo red: {ex.Message}");
            }
            return sb.ToString();
        }

        public string ObtenerInformacionUsuario()
        {
            try
            {
                string usuario = Environment.UserName;
                string dominio = Environment.UserDomainName;
                string host = Environment.MachineName;
                return $"Usuario: {usuario} | Dominio: {dominio} | Host: {host}";
            }
            catch (Exception ex)
            {
                return $"Error obteniendo usuario: {ex.Message}";
            }
        }

        public void ExportarInformacion(string info)
        {
            try
            {
                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "InfoPC.txt");
                File.WriteAllText(path, info);
                MessageBox.Show($"Información exportada a: {path}", "Exportación", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exportando información: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string ObtenerMarcaModeloPC()
        {
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string marca = obj["Manufacturer"]?.ToString() ?? "Desconocido";
                    string modelo = obj["Model"]?.ToString() ?? "Desconocido";
                    return $"Marca: {marca} | Modelo: {modelo}";
                }
            }
            catch (Exception ex)
            {
                return $"Error obteniendo marca/modelo: {ex.Message}";
            }
            return "Marca/Modelo no disponible";
        }

        public string ObtenerInformacionCompleta()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(ObtenerMarcaModeloPC());
            sb.AppendLine("Motherboard: " + ObtenerInfoMotherboard());
            sb.AppendLine("Tarjeta Gráfica: " + ObtenerInformacionGrafica());
            sb.AppendLine("RAM: " + ObtenerInformacionTipoRAM());
            sb.AppendLine("Cantidad RAM: " + ObtenerInformacionCantidadRAM());
            sb.AppendLine("Cantidad Slots: " + ObtenerInformacionCantidadSlots());
            sb.AppendLine("Procesador: " + ObtenerInformacionProcesador());
            sb.AppendLine("SO: " + ObtenerInformacionSO());
            sb.AppendLine("Usuario: " + ObtenerInformacionUsuario());
            sb.AppendLine("Red:");
            sb.AppendLine(ObtenerInformacionRed());
            sb.AppendLine("Almacenamiento:");
            DriveInfo[] allDrives = DriveInfo.GetDrives();
            foreach (DriveInfo d in allDrives)
            {
                if (d.IsReady)
                {
                    sb.AppendLine($"Disco: {d.Name} | Tipo: {ObtenerInformacionTipoDisco(d)} | Formato: {d.DriveFormat} | Espacio Disponible: {FormatBytes(d.TotalFreeSpace)} | Espacio Utilizado: {FormatBytes(d.TotalSize - d.TotalFreeSpace)} | Espacio Total: {FormatBytes(d.TotalSize)}");
                }
            }
            return sb.ToString();
        }

        private void btnExportarInfo_Click(object sender, EventArgs e)
        {
            string info = ObtenerInformacionCompleta();
            ExportarInformacion(info);
        }

    }
}
