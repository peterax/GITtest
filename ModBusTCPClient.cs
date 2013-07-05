// /////////////////////////////////////////////////////////////////////////////////////////////////
// // Classname       : ModbusTcpClient
// // Author          : Claes Park
// // Original date   : 2011-10-12
// // Description     : Base class for ModbusTCP Client
// //
// // Revision history (date of change and short description):
// // 2011-10-12: Original release
// // 
// // 
// ///////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tiara.Network;
using Tiara.Logging;

namespace Tiara.IO
{
    public class ModBusTCPClient : Modbus
    {
        private Tiara.Network.Modbus.ConnectionStatus _connectionStatus;
        private Hashtable _lockTableIORequest;
        private int transactionId;

        public enum enuOnOff : int
        {
            SetOn = 1,
            SetOff = 0,
        }

        Tiara.Network.ModbusTCP.Master master;      //The modbus object handling the communication

        string _serverIP = System.String.Empty;
        int _serverPort = 502;  //Default

        int _updateRate = 500;  //default 500 ms

       // public ModBusTCPClient(string serverIP, int serverPort)
        public ModBusTCPClient(ModbusTcpIp _modbusT)
            : base()
        {
            _lockTableIORequest = new Hashtable();
        }


        public override ConnectionStatus Status
        {
            get
            {
                return (_connectionStatus);
            }
            //set
            //{
            //    _status = value;
            //}
        }


        public string ServerIP
        {
            get { return _serverIP; }
            set { _serverIP = value; }
        }

        public int ServerPort
        {
            get { return _serverPort; }
            set { _serverPort = value; }
        }

        public bool Start()
        {
            bool rtv = true;

            if (master != null)
                master.Dispose();

            master = new Tiara.Network.ModbusTCP.Master();

            return rtv;
        }

        //Pax variant
        public bool Start(ModbusTcpIp _tcpipClient)
        {
            bool rtv = true;

            if (master != null)
                master.Dispose();

              //////          master = new Tiara.Network.ModbusTCP.Master(_serverIP, _serverPort);
            master = new Tiara.Network.ModbusTCP.Master(_tcpipClient);

            return rtv;
        }

        public bool Stop()
        {
            bool rtv = true;
            //master.disconnect();
            //master.Dispose();

            return rtv;
        }

        public bool Reset()
        {
            bool rtv = true;
            Stop();
            rtv = Start();
            return rtv;
        }
        public bool ReadIOTag(string tagAddress, bool isInput, out enuOnOff value)
        {
            int StartAddress = 0;
            byte Length = 1;
            byte[] result = null;
            Tiara.IO.ModBusTCPClient.enuOnOff onOff;
            bool rtv = false;

            try
            {
                if (master.connected)//PAx prylar
                {
                    int ID = ++transactionId;

                    // check if address is word.bit format
                    if (tagAddress.Contains('.'))
                    {
                        int bit = 0;
                        int word = 0;
                        int.TryParse(tagAddress.Substring(tagAddress.IndexOf('.', 0) + 1), out bit);
                        int.TryParse(tagAddress.Substring(0, tagAddress.IndexOf('.')), out word);
                        StartAddress = (word) * 16 + (15 - bit);
                    }
                    else
                    {
                        StartAddress = System.Convert.ToInt32(tagAddress);
                    }
                    //result = new byte[1];
                    if (isInput)
                        master.ReadDiscreteInputs(ID, StartAddress, Length, ref result);
                    else
                        master.ReadCoils(ID, StartAddress, Length, ref result);

                }


            }
            catch (Exception ex)
            {
                ////////////////// Stop();KAnske så
                //started = false;
                rtv = false;
                //Added 2011-01-25, PaCl
                throw new Exception(new StringBuilder("Error reading io tag ").Append(tagAddress).Append(": ").Append(ex.Message).ToString());
            }
            if (result != null && result.Length > 0)
            {
                // Maska ut LSB ur svaret (ADAM returnerade 8 bitar
                // trots att man bara läste en).
                int iResultLsb = (int)result[0] & 1;
                value = (enuOnOff)((iResultLsb == 1) ? 1 : 0);
                rtv = true;
            }
            else
                value = (enuOnOff)0;

            return rtv;
        }


        public bool WriteIOTag(string tagAddress, enuOnOff value)
        {
            bool rtv = true;
            string res;
            int rtc;
            byte[] result = null;

            //if (!master.connected)
            //    Start();

            try
            {
                if (master.connected)//PAx prylar
                {
                    bool onOff = value == enuOnOff.SetOn;
//                    byte[] result = null;

                    int ID = ++transactionId;
                    int StartAddress = 0;

                    // check if address is word.bit format
                    if (tagAddress.Contains('.'))
                    {
                        int bit = 0;
                        int word = 0;
                        int.TryParse(tagAddress.Substring(tagAddress.IndexOf('.', 0) + 1), out bit);
                        int.TryParse(tagAddress.Substring(0, tagAddress.IndexOf('.')), out word);
                        StartAddress = (word) * 16 + (15 - bit);
                    }
                    else
                    {
                        StartAddress = System.Convert.ToInt32(tagAddress);
                    }
                    master.WriteSingleCoils(ID, StartAddress, onOff, ref result);
                }
            }
            catch (Exception ex)
            {
                Stop();
                //started = false;
                rtv = false;
                //Added 2011-01-25, PaCl
                throw new Exception(new StringBuilder("Error writing io tag ").Append(tagAddress).Append(": ").Append(ex.Message).ToString());
            }
            rtv = (result != null) && (result.Length > 0);
            return rtv;
        }

        //Added 2011-08-15, PaCl
        /// <summary>
        /// Read value (as string) from ModbusTCP-tag (int, string etc)
        /// </summary>
        public bool ReadDataTag(string tagAddress, string dataType, bool isInput, out string value)
        {
            bool rtv = false;
            string res;
            int rtc;
            byte[] responseData = null;
            string convertedValue = "";
            value = System.String.Empty;

            int StartAddress = 0;
            byte Length = 1;
            byte[] result = null;

            bool goodResponse = false;
            bool tagIsArray = false;
            int startBracket = 0;
            int endBracket = 0;
            int noOfTags = 1;

            try
            {
                if (master.connected)
                {
                    int ID = ++transactionId;

                    bool adressResolved = ResolveTagAddress(tagAddress, dataType, out StartAddress, out tagIsArray, out noOfTags);

                    switch (dataType)
                    {
                        case "Bit":
                            // Number of COILS to read for EACH tag
                            Length = 1;

                            // Multiply w noOfTags to get noOfCoils to request
                            Length = (byte)(Length * noOfTags);

                            // Read (Length) inputs/coils starting at (StartAddress)
                            if (isInput) 
                                master.ReadDiscreteInputs(ID, StartAddress, Length, ref responseData);
                            else
                                master.ReadCoils(ID, StartAddress, Length, ref responseData);

                            // 8 coils per byte in response
                            if (responseData != null && responseData.Length == ((Length + 7) / 8)) //OK response data?
                            {
                                for (int i = 0; i < noOfTags; i++)
                                {
                                    int byteIndex = i / 8;
                                    int bitIndex = i % 8;
                                    // make mask
                                    int bitmask = (int)Math.Pow(2, bitIndex);
                                    // mask
                                    int tagvalue = (int)responseData[byteIndex] & bitmask;
                                    // evaluate masking
                                    if (tagvalue != 0)
                                        tagvalue = 1;
                                    else
                                        tagvalue = 0;

                                    // Add to result string separated with #
                                    convertedValue = convertedValue + Convert.ToString(tagvalue) + "#";
                                }
                                // Clean up # if single value, else add ¤
                                if (noOfTags == 1)
                                    convertedValue = convertedValue.TrimEnd('#');
                                else
                                    convertedValue = convertedValue + "¤";

                                goodResponse = true;
                            }
                            break;

                        case "Int16":
                            // Number of COILS to read for EACH tag
                            Length = 1;
                            // Multiply w noOfTags to get noOfCoils to request
                            Length = (byte)(Length * noOfTags);

                            responseData = new byte[Length * 2];

                            master.ReadInputRegister(ID, StartAddress, Length, ref responseData);

                            if (responseData != null && responseData.Length == Length * 2)   //OK response data?
                                for (int i = 0; i < noOfTags; i++)
                                {
                                    convertedValue = convertedValue + 
                                        Convert.ToString(
                                        responseData[i*2] * Math.Pow(2, 8) +
                                        responseData[i*2+1]);

                                    convertedValue = convertedValue + "#";
                                }
                            // Clean up # if single value, else add ¤
                            if (noOfTags == 1)
                                convertedValue = convertedValue.TrimEnd('#');
                            else
                                convertedValue = convertedValue + "¤";

                            goodResponse = true;
                            break;

                        case "Int32":
                            // Number of COILS to read for EACH tag
                            Length = 2;

                            // Multiply w noOfTags to get noOfCoils to request
                            Length = (byte)(Length * noOfTags);

                            responseData = new byte[Length * 2];

                            master.ReadInputRegister(ID, StartAddress, Length, ref responseData);

                            if (responseData != null && responseData.Length == Length * 2)   //OK response data?
                                for (int i = 0; i < noOfTags; i++)
                                {
                                    convertedValue = convertedValue +
                                        Convert.ToString(
                                            responseData[i * 4 + 0] * Math.Pow(2, 24) +
                                            responseData[i * 4 + 1] * Math.Pow(2, 16) +
                                            responseData[i * 4 + 2] * Math.Pow(2, 8) +
                                            responseData[i * 4 + 3]);

                                    convertedValue = convertedValue + "#";
                                }
                            // Clean up # if single value, else add ¤
                            if (noOfTags == 1)
                                convertedValue = convertedValue.TrimEnd('#');
                            else
                                convertedValue = convertedValue + "¤";

                            goodResponse = true;
                            break;

                        case "Int64":
                            // Number of COILS to read for EACH tag
                            Length = 4;

                            // Multiply w noOfTags to get noOfCoils to request
                            Length = (byte)(Length * noOfTags);

                            responseData = new byte[Length * 2];

                            master.ReadInputRegister(ID, StartAddress, Length, ref responseData);

                            if (responseData != null && responseData.Length == Length * 2)   //OK response data?
                                for (int i = 0; i < noOfTags; i++)
                                {
                                    convertedValue = convertedValue +
                                        Convert.ToString(
                                            responseData[i * 8 + 0] * Math.Pow(2, 56) +
                                            responseData[i * 8 + 1] * Math.Pow(2, 48) +
                                            responseData[i * 8 + 2] * Math.Pow(2, 40) +
                                            responseData[i * 8 + 3] * Math.Pow(2, 32) +
                                            responseData[i * 8 + 4] * Math.Pow(2, 24) +
                                            responseData[i * 8 + 5] * Math.Pow(2, 16) +
                                            responseData[i * 8 + 6] * Math.Pow(2, 8) +
                                            responseData[i * 8 + 7]);

                                    convertedValue = convertedValue + "#";
                                }
                            // Clean up # if single value, else add ¤
                            if (noOfTags == 1)
                                convertedValue = convertedValue.TrimEnd('#');
                            else
                                convertedValue = convertedValue + "¤";

                            goodResponse = true;
                            break;

                        case "String":
                            // check if address is word[length] format
                            startBracket = 0;
                            endBracket = 0;
                            int tagLength = 0;

                            startBracket = tagAddress.IndexOf('[');
                            endBracket = tagAddress.IndexOf(']');
                            int.TryParse(tagAddress.Substring(startBracket + 1, endBracket - startBracket - 1), out tagLength);
                            int.TryParse(tagAddress.Substring(0, startBracket), out StartAddress);

                            // ToDo: Udda taglängd då???
                            Length = (byte)(tagLength / 2);
                            responseData = new byte[tagLength];

                            result = null;

                            master.ReadInputRegister(ID, StartAddress, Length, ref responseData);

                            if (responseData != null && responseData.Length == Length * 2)   //Bad response data?
                            {
                                convertedValue = Tiara.Utilities.Converting.ByteArrayToString(responseData);
                                if (convertedValue.Contains("\0"))
                                    //convertedValue = "";
                                    convertedValue = convertedValue.Remove(convertedValue.IndexOf("\0"));
                            }
                            break;

                    }

                    if ((dataType != "Bit" && responseData != null && responseData.Length == Length * 2) ||
                        (dataType == "Bit" && responseData != null && responseData.Length == Length) ||//Good response data?
                        (goodResponse == true))
                    {
                        //Convert response data
                        value = convertedValue;
                        rtv = true;
                    }
                }
            }
            catch (Exception ex)
            {
                /////////////Stop();
                //started = false;
                rtv = false;
                //Added 2011-01-25, PaCl
                throw new Exception(new StringBuilder("Error reading data tag ").Append(tagAddress).Append(": ").Append(ex.Message).ToString());
            }
            return rtv;
        
        }

        //Added 2011-08-15, PaCl
        /// <summary>
        /// Write value to OPC-tag (int, string etc)
        /// </summary>
        public bool WriteDataTag(string tagAddress, string dataType, string value)
        {
            bool rtv = true;
            string res;
            int rtc;
            byte[] result = null;

            //if (!master.connected)
            //    Start();

            try
            {
                if (master.connected)//PAx prylar
                {
                    int ID = ++transactionId;
                    switch (dataType)
                    {
                        case "Bit":
                            bool retval;

                            if (value == "0")
                                retval = WriteIOTag(tagAddress, Tiara.IO.ModBusTCPClient.enuOnOff.SetOff);
                            else
                                retval = WriteIOTag(tagAddress, Tiara.IO.ModBusTCPClient.enuOnOff.SetOn);

                            break;

                        case "Int16":
                            int StartAddress = System.Convert.ToInt32(tagAddress);
                            byte Length = 1;

                            int valueInt;
                            if (!int.TryParse(value.Trim(), out valueInt))
                                return false;

                            byte[] valueBytes = new byte[2];
                            result = null;

                            valueBytes[0] = (byte)(valueInt / Math.Pow(2, 8));
                            valueBytes[1] = (byte)(valueInt % Math.Pow(2, 8));

                            master.WriteMultipleRegister(ID, StartAddress, Length, valueBytes, ref result);

                            break;

                        case "Int32":
                            StartAddress = System.Convert.ToInt32(tagAddress);
                            Length = 2;

                            if (!int.TryParse(value.Trim(), out valueInt))
                                return false;

                            valueBytes = new byte[4];
                            result = null;

                            valueBytes[0] = (byte)(valueInt / Math.Pow(2, 24));
                            valueBytes[1] = (byte)((valueInt % Math.Pow(2, 24)) / Math.Pow(2, 16));
                            valueBytes[2] = (byte)((valueInt % Math.Pow(2, 16)) / Math.Pow(2, 8));
                            valueBytes[3] = (byte)(valueInt % Math.Pow(2, 8));

                            master.WriteMultipleRegister(ID, StartAddress, Length, valueBytes, ref result);

                            break;

                        case "Int64":
                            StartAddress = System.Convert.ToInt32(tagAddress);
                            Length = 4;

                            if (!int.TryParse(value.Trim(), out valueInt))
                                return false;

                            valueBytes = new byte[8];
                            result = null;

                            valueBytes[0] = (byte)(valueInt / Math.Pow(2, 56));
                            valueBytes[1] = (byte)((valueInt % Math.Pow(2, 56)) / Math.Pow(2, 48));
                            valueBytes[2] = (byte)((valueInt % Math.Pow(2, 48)) / Math.Pow(2, 40));
                            valueBytes[3] = (byte)((valueInt % Math.Pow(2, 40)) / Math.Pow(2, 32));
                            valueBytes[4] = (byte)((valueInt % Math.Pow(2, 32)) / Math.Pow(2, 24));
                            valueBytes[5] = (byte)((valueInt % Math.Pow(2, 24)) / Math.Pow(2, 16));
                            valueBytes[6] = (byte)((valueInt % Math.Pow(2, 16)) / Math.Pow(2, 8));
                            valueBytes[7] = (byte)(valueInt % Math.Pow(2, 8));

                            master.WriteMultipleRegister(ID, StartAddress, Length, valueBytes, ref result);

                            break;

                        case "String":
                            // check if address is word[length] format
                            int startBracket = 0;
                            int endBracket = 0;
                            int tagLength = 0;

                            startBracket = tagAddress.IndexOf('[');
                            endBracket = tagAddress.IndexOf(']');
                            int.TryParse(tagAddress.Substring(startBracket + 1, endBracket - startBracket - 1), out tagLength);
                            int.TryParse(tagAddress.Substring(0, startBracket), out StartAddress);

                            // Pad data to match tagLength
                            value = value.PadRight(tagLength, ' ');

                            Length = (byte)(tagLength / 2);
                            result = null;
                            valueBytes = new byte[tagLength];
                            valueBytes = Tiara.Utilities.Converting.ConvertStringToByte(value.Substring(0, tagLength));

                            master.WriteMultipleRegister(ID, StartAddress, Length, valueBytes, ref result);

                            break;

                    }
                }
            }
            catch (Exception ex)
            {
              //  Stop(); Tycker inte PAx
                rtv = false;
                //Added 2011-01-25, PaCl
                throw new Exception(new StringBuilder("Error writing data tag ").Append(tagAddress).Append(": ").Append(ex.Message).ToString());
            }
            rtv = (result != null && result.Length > 0);

            return rtv;
        }

        private object GetOrAddIOHandleLock(string serverName)
        {
            object result = _lockTableIORequest[serverName];

            if (result == null)
            {
                result = new object();
                lock (_lockTableIORequest.SyncRoot)
                {
                    _lockTableIORequest[serverName] = result;
                }
            }

            return result;
        }

        //----------------------------------------------------------------------
        // Resolve address formats to Modbus standard
        // Accepted formats:
        // 2 : Needs no resolving, use as it is for register or coils
        // 2.9 : Converts to coil no (coil 0 is MSB of register so 2.9 = coil 39)
        // 
        // Siemens S7: Address format that match the PLC addresses on the other side of the Anybus:
        // s7i2: Byte based Siemens address for registers: Byte 2 eq register 1, 4 eq 2 etc)
        // s7i4.1: Siemens byte.bit format: s7i4.1 = bit1 of byte 4 (high byte of reg 1) = bit9 of reg2 = coil 38)
        // s7q: Adds offset of 1024 for registers and 16384 for coils (made for Anybus)
        //
        // All tags can be addressed as arrays by adding [xx]. Result values are then returned in
        // CodeIT table format (###¤).
        //

        // Nota Bene: Siemensadressering är still under progress, därav bortkommenteringarna!
        //
        public bool ResolveTagAddress(string tagAddress, string dataType, out int StartAddress, out bool tagIsArray, out int noOfTags)
        {
            bool rtv = true;

            StartAddress = 0;
            tagIsArray = false;
            noOfTags = 1;

            // Check if address is formatted 'address[length]' and extract noOfTags and address
            tagIsArray = (tagAddress.Contains('[')) && (tagAddress.Contains(']'));
            if (tagIsArray)
            {
                int startBracket = tagAddress.IndexOf('[');
                int endBracket = tagAddress.IndexOf(']');
                int.TryParse(tagAddress.Substring(startBracket + 1, endBracket - startBracket - 1), out noOfTags);
                tagAddress = tagAddress.Substring(0, startBracket);
            }

            // Check if address is in Siemens S7 byte.bit format (prefix s7i or s7q)
            //
            // First, convert to lower case, just in case
            tagAddress = tagAddress.ToLower();
            if ((tagAddress.Contains("s7")) && (tagAddress.Contains('.'))) //Chg 2012-12-10 PAx
            {
       /*         int bit = 0;
                int byteNo = 0;
                int.TryParse(tagAddress.Substring(tagAddress.IndexOf('.', 0) + 1), out bit);
                int.TryParse(tagAddress.Substring(3, tagAddress.IndexOf('.') - 3), out byteNo);
                // input address, no offset
                if ((tagAddress.Substring(0, 3) == "s7i"))
                    StartAddress = (byteNo) * 8 + (7 - bit);
                // output address, offset 16384 in Anybus coil area
                if ((tagAddress.Substring(0, 3) == "s7q"))
                    StartAddress = (byteNo) * 8 + (7 - bit) + 16384;
*/
            }
            // check if address is Non Siemens word.bit format
            else if ((!tagAddress.Contains("s7")) && (tagAddress.Contains('.'))) //Chg 2012-12-10 PAx
            {
/*                int bit = 0;
                int word = 0;
                int.TryParse(tagAddress.Substring(tagAddress.IndexOf('.', 0) + 1), out bit);
                int.TryParse(tagAddress.Substring(0, tagAddress.IndexOf('.')), out word);
                StartAddress = (word) * 16 + (15 - bit);
 */
            }
            // check if address is in Siemens S7 format (prefix S7I or S7Q) w o dot notation
            // this is presumed to be a byte based start address:
            // byte 0 means start register 0 (consisting of byte 0 and 1),
            // byte 2 means start register 1 (consisting of byte 2 and 3)
            else if ((tagAddress.Contains("s7")) && (!tagAddress.Contains('.'))) //Chg 2012-12-10 PAx
            {
/*                int byteNo = 0;
                int.TryParse(tagAddress.Remove(0, 3), out byteNo);
                //word =
                // input address, no offset
                if ((tagAddress.Substring(0, 3) == "s7i"))
                    StartAddress = byteNo / 2;
                // output address, offset 1024 in Anybus register area
                if ((tagAddress.Substring(0, 3) == "s7q"))
                    StartAddress = byteNo / 2 + 1024;
 */
            }
            else
            {
                StartAddress = System.Convert.ToInt32(tagAddress);
            }
            return rtv;
        }


    }
}
