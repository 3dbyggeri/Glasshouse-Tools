#region copyright notice
/*
Original work Copyright(c) 2018-2021 COWI
    
Copyright © COWI and individual contributors. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

    1) Redistributions of source code must retain the above copyright notice,
    this list of conditions and the following disclaimer.

    2) Redistributions in binary form must reproduce the above copyright notice,
    this list of conditions and the following disclaimer in the documentation
    and/or other materials provided with the distribution.

    3) Neither the name of COWI nor the names of its contributors may be used
    to endorse or promote products derived from this software without specific
    prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS “AS IS”
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
THE POSSIBILITY OF SUCH DAMAGE.

GlasshouseExcel may utilize certain third party software. Such third party software is copyrighted by their respective owners as indicated below.
Netoffice - MIT License - https://github.com/NetOfficeFw/NetOffice/blob/develop/LICENSE.txt
Excel DNA - zlib License - https://github.com/Excel-DNA/ExcelDna/blob/master/LICENSE.txt
RestSharp - Apache License - https://github.com/restsharp/RestSharp/blob/develop/LICENSE.txt
Newtonsoft - The MIT License (MIT) - https://github.com/JamesNK/Newtonsoft.Json/blob/master/LICENSE.md
*/
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Reflection;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using RTD.Server;
using Application = NetOffice.ExcelApi.Application;

namespace GlasshouseExcel
{
    public class clsProgram : IExcelAddIn, IRTDClient
    {
        static object _this = null;

        // connection to server
        static IRTDServer _server;
        private static bool connected = false;

        // never called for addins!
        public void AutoClose()
        {
            throw new NotImplementedException();
            // _server.UnRegister();
        }

        public void AutoOpen()
        {
            _this = this;
            ConnectToSonnect();
        }


        public static void ConnectToSonnect()
        {
            if (connected)
            {
                return;
            }

            try
            {
                // setup connection to server
                var a = new EndpointAddress("net.tcp://localhost:8080/hello");
                var b = new NetTcpBinding();
                //var f = new DuplexChannelFactory<IRTDServer>(this, b, a);
                var f = new DuplexChannelFactory<IRTDServer>(_this, b, a);
                _server = f.CreateChannel();
                _server.Register(); // this makes the server get a callback channel to us so it can call SendValue

                // increase RTD refresh rate since the 2 seconds default is too slow (move as setting to ribbon later)
                object app;
                object rtd;
                app = ExcelDnaUtil.Application;
                rtd = app.GetType().InvokeMember("RTD", BindingFlags.GetProperty, null, app, null);
                rtd.GetType().InvokeMember("ThrottleInterval", BindingFlags.SetProperty, null, rtd, new object[] { 100 });
                connected = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        // event boilerplate stuff
        public delegate void ValueSentHandler(ValueSentEventArgs args);
        public static event ValueSentHandler OnValueSent;
        public class ValueSentEventArgs : EventArgs
        {
            public double Value { get; private set; }
            public ValueSentEventArgs(double Value)
            {
                this.Value = Value;
            }
        }

        // this gets called by the server if there is a new value
        public void SendValue(double x)
        {

            // invert method call from WCF into event for Rx
            if (OnValueSent != null)
                OnValueSent(new ValueSentEventArgs(x));
        }

        // this function accepts an observer and calls OnNext on it every time the OnValueSent event is raised
        // this is the part that chains individual events into streams
        // it returns an IDisposable that can be called after the stream is complete 
        static Func<IObserver<double>, IDisposable> Event2Observable = observer =>
        {
            OnValueSent += d => observer.OnNext(d.Value);   // when a new value was sent, call OnNext on the Observers.
            return Disposable.Empty;                        // we'll keep listening until Excel is closed, there's no end to this stream
        };

        // this calls the wrapper to actually put it all together
        [ExcelFunction("Gets realtime values from server")]
        public static object GetValues()
        {

            // a delegate that creates an observable over Event2Observable
            Func<IObservable<double>> f2 = () => Observable.Create<double>(Event2Observable);

            //  pass that to Excel wrapper   
            return RxExcel.Observe("GetValues", null, f2);
        }
    }
}
