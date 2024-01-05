using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace Tatais.DMX
{

    /// <summary>
    /// Dataset ofn state Unit, Channels at time stamp
    /// </summary>
    [XmlRoot(ElementName ="DMXSet")]
    public class DMXData : IComparable<DMXData>
    {
        public int CompareTo(DMXData other)
        {
            if ((other == null) || (!(other is DMXData))) return 1;
            else
                return Comparer<int>.Default.Compare(this.Order, other.Order);
        }
        [XmlElement(ElementName ="Order")]
        public int Order { get; set; }
        [XmlElement(ElementName = "Unit")]
        public int Unit { get; set; }
        [XmlElement(ElementName = "Red")]
        public int Red { get; set; }
        [XmlElement(ElementName = "Green")]
        public int Green { get; set; }
        [XmlElement(ElementName = "Blue")]
        public int Blue { get; set; }
        [XmlElement(ElementName = "White")]
        public int White { get; set; }
        [XmlElement(ElementName = "Zoom")]
        public int Zoom { get; set; }
        [XmlElement(ElementName = "Time")]
        public int Sec { get; set; }
        /// <summary>
        /// Set Unit state at time
        /// </summary>
        /// <param name="sec">Time stamp</param>
        /// <param name="unit">Unit num</param>
        /// <param name="red">Red chanel data</param>
        /// <param name="green">Green chanel data</param>
        /// <param name="blue">Blue chanel data</param>
        /// <param name="white">White chanel data</param>
        /// <param name="zoom">Zoom chanel data</param>
        public DMXData(int order, int sec, int unit, int red, int green, int blue, int white, int zoom)
        {
            this.Unit = unit;
            this.Red = red;
            this.Green = green;
            this.Blue = blue;
            this.White = white;
            this.Zoom = zoom;
            this.Sec = sec;
            this.Order = order;
        }
        public DMXData() { }
    }
    /*[XmlRoot(ElementName="AutomationList")]
    public class AutomationList
    {
        [XmlElement(ElementName ="DMXSet")]
        public List<DMXData> DMXData { get; set; }
    }*/
    internal class Automation
    {
        
        public List<DMXData> automationListDMXData;
        /// <summary>
        /// Auto switch channels data save to file sample
        /// </summary>
        public void SaveData(string filename)
       {
            
            automationListDMXData = new List <DMXData> ();
            //Dummu data
            automationListDMXData.Add(new DMXData(1, 3000, 1, 255, 0, 255, 0, 100));
            automationListDMXData.Add(new DMXData(3,1100,1,255,255,0,100,0));
            automationListDMXData.Add(new DMXData(2,20000, 1, 0, 200, 100, 100,50));
            automationListDMXData.Add(new DMXData(4,110000, 1, 0, 200, 100, 100, 50));
            //Save XML file
            XmlSerializer serializer = new XmlSerializer(typeof (List<DMXData>));
            TextWriter writer = new StreamWriter(filename);
            serializer.Serialize(writer, automationListDMXData);
            writer.Close();

        }
        /// <summary>
        /// Load data from xml file
        /// </summary>
        /// <param name="filename">file to load</param>
        public void LoadData(string filename)
        {
            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(List<DMXData>));
                automationListDMXData = new List<DMXData>();

                using (FileStream fs = File.OpenRead(filename))
                {
                    automationListDMXData = (List<DMXData>)xmlSerializer.Deserialize(fs);
                    MessageBox.Show("File loaded","Loading Program Instruction");
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message,"XML format error", MessageBoxButtons.OK);
            }
        }
    }
}
