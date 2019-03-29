
    using System.Diagnostics;
    using System.Web.Services;
    using System.ComponentModel;
    using System.Web.Services.Protocols;
    using System;
    using System.Xml.Serialization;
    
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="cmcc_mas_wbsSoapBinding", Namespace="http://www.csapi.org/service")]
    public partial class cmcc_mas_wbs : System.Web.Services.Protocols.SoapHttpClientProtocol {
        
        private System.Threading.SendOrPostCallback APRegistrationOperationCompleted;
        
        private System.Threading.SendOrPostCallback APStatusRepOperationCompleted;
        
        private System.Threading.SendOrPostCallback APLogOutOperationCompleted;
        
        private System.Threading.SendOrPostCallback PauseAPOperationCompleted;
        
        private System.Threading.SendOrPostCallback RecoveryAPOperationCompleted;
        
        private System.Threading.SendOrPostCallback APSvcAuthenticOperationCompleted;
        
        private System.Threading.SendOrPostCallback APSvcPerfCmdOperationCompleted;
        
        private System.Threading.SendOrPostCallback APSvcAlarmOperationCompleted;
        
        private System.Threading.SendOrPostCallback APSvcPerfReportOperationCompleted;
        
        private System.Threading.SendOrPostCallback startNotificationOperationCompleted;
        
        private System.Threading.SendOrPostCallback stopNotificationOperationCompleted;
        
        private System.Threading.SendOrPostCallback sendSmsOperationCompleted;
        
        private System.Threading.SendOrPostCallback notifySmsDeliveryStatusOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetReceivedSmsOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetSmsDeliveryStatusOperationCompleted;
        
        private System.Threading.SendOrPostCallback notifySmsReceptionOperationCompleted;
        
        private System.Threading.SendOrPostCallback sendMessageOperationCompleted;
        
        private System.Threading.SendOrPostCallback getMessageDeliveryStatusOperationCompleted;
        
        private System.Threading.SendOrPostCallback getReceivedMessagesOperationCompleted;
        
        private System.Threading.SendOrPostCallback getMessageOperationCompleted;
        
        private System.Threading.SendOrPostCallback notifyMessageReceptionOperationCompleted;
        
        private System.Threading.SendOrPostCallback notifyMessageDeliveryReceiptOperationCompleted;
        
        private System.Threading.SendOrPostCallback getLocationOperationCompleted;
        
        private System.Threading.SendOrPostCallback getLocationForGroupOperationCompleted;
        
        private System.Threading.SendOrPostCallback startPeriodicNotificationOperationCompleted;
        
        private System.Threading.SendOrPostCallback EndNotificationOperationCompleted;
        
        private System.Threading.SendOrPostCallback LocationNotificationOperationCompleted;
        
        private System.Threading.SendOrPostCallback LocationErrorOperationCompleted;
        
        private System.Threading.SendOrPostCallback LocationEndOperationCompleted;
        
        private System.Threading.SendOrPostCallback sendPushOperationCompleted;
        
        private System.Threading.SendOrPostCallback getPushDeliveryStatusOperationCompleted;
        
        private System.Threading.SendOrPostCallback notifyPushDeliveryReceiptOperationCompleted;
        
        private System.Threading.SendOrPostCallback makeUssdOperationCompleted;
        
        private System.Threading.SendOrPostCallback handleUssdOperationCompleted;
        
        private System.Threading.SendOrPostCallback ussdContinueOperationCompleted;
        
        private System.Threading.SendOrPostCallback endUssdOperationCompleted;
        
        private System.Threading.SendOrPostCallback notifyUssdEndOperationCompleted;
        
        /// <remarks/>
        public cmcc_mas_wbs() {
            this.Url = "http://127.0.0.1/services/cmcc_mas_wbs";
        }
        
        /// <remarks/>
        public event APRegistrationCompletedEventHandler APRegistrationCompleted;
        
        /// <remarks/>
        public event APStatusRepCompletedEventHandler APStatusRepCompleted;
        
        /// <remarks/>
        public event APLogOutCompletedEventHandler APLogOutCompleted;
        
        /// <remarks/>
        public event PauseAPCompletedEventHandler PauseAPCompleted;
        
        /// <remarks/>
        public event RecoveryAPCompletedEventHandler RecoveryAPCompleted;
        
        /// <remarks/>
        public event APSvcAuthenticCompletedEventHandler APSvcAuthenticCompleted;
        
        /// <remarks/>
        public event APSvcPerfCmdCompletedEventHandler APSvcPerfCmdCompleted;
        
        /// <remarks/>
        public event APSvcAlarmCompletedEventHandler APSvcAlarmCompleted;
        
        /// <remarks/>
        public event APSvcPerfReportCompletedEventHandler APSvcPerfReportCompleted;
        
        /// <remarks/>
        public event startNotificationCompletedEventHandler startNotificationCompleted;
        
        /// <remarks/>
        public event stopNotificationCompletedEventHandler stopNotificationCompleted;
        
        /// <remarks/>
        public event sendSmsCompletedEventHandler sendSmsCompleted;
        
        /// <remarks/>
        public event notifySmsDeliveryStatusCompletedEventHandler notifySmsDeliveryStatusCompleted;
        
        /// <remarks/>
        public event GetReceivedSmsCompletedEventHandler GetReceivedSmsCompleted;
        
        /// <remarks/>
        public event GetSmsDeliveryStatusCompletedEventHandler GetSmsDeliveryStatusCompleted;
        
        /// <remarks/>
        public event notifySmsReceptionCompletedEventHandler notifySmsReceptionCompleted;
        
        /// <remarks/>
        public event sendMessageCompletedEventHandler sendMessageCompleted;
        
        /// <remarks/>
        public event getMessageDeliveryStatusCompletedEventHandler getMessageDeliveryStatusCompleted;
        
        /// <remarks/>
        public event getReceivedMessagesCompletedEventHandler getReceivedMessagesCompleted;
        
        /// <remarks/>
        public event getMessageCompletedEventHandler getMessageCompleted;
        
        /// <remarks/>
        public event notifyMessageReceptionCompletedEventHandler notifyMessageReceptionCompleted;
        
        /// <remarks/>
        public event notifyMessageDeliveryReceiptCompletedEventHandler notifyMessageDeliveryReceiptCompleted;
        
        /// <remarks/>
        public event getLocationCompletedEventHandler getLocationCompleted;
        
        /// <remarks/>
        public event getLocationForGroupCompletedEventHandler getLocationForGroupCompleted;
        
        /// <remarks/>
        public event startPeriodicNotificationCompletedEventHandler startPeriodicNotificationCompleted;
        
        /// <remarks/>
        public event EndNotificationCompletedEventHandler EndNotificationCompleted;
        
        /// <remarks/>
        public event LocationNotificationCompletedEventHandler LocationNotificationCompleted;
        
        /// <remarks/>
        public event LocationErrorCompletedEventHandler LocationErrorCompleted;
        
        /// <remarks/>
        public event LocationEndCompletedEventHandler LocationEndCompleted;
        
        /// <remarks/>
        public event sendPushCompletedEventHandler sendPushCompleted;
        
        /// <remarks/>
        public event getPushDeliveryStatusCompletedEventHandler getPushDeliveryStatusCompleted;
        
        /// <remarks/>
        public event notifyPushDeliveryReceiptCompletedEventHandler notifyPushDeliveryReceiptCompleted;
        
        /// <remarks/>
        public event makeUssdCompletedEventHandler makeUssdCompleted;
        
        /// <remarks/>
        public event handleUssdCompletedEventHandler handleUssdCompleted;
        
        /// <remarks/>
        public event ussdContinueCompletedEventHandler ussdContinueCompleted;
        
        /// <remarks/>
        public event endUssdCompletedEventHandler endUssdCompleted;
        
        /// <remarks/>
        public event notifyUssdEndCompletedEventHandler notifyUssdEndCompleted;
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/APRegistration", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("APRegistrationRsp", Namespace="http://www.csapi.org/schema/ap")]
        public APRegistrationRsp APRegistration([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ap")] APRegistrationReq APRegistrationReq) {
            object[] results = this.Invoke("APRegistration", new object[] {
                        APRegistrationReq});
            return ((APRegistrationRsp)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginAPRegistration(APRegistrationReq APRegistrationReq, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("APRegistration", new object[] {
                        APRegistrationReq}, callback, asyncState);
        }
        
        /// <remarks/>
        public APRegistrationRsp EndAPRegistration(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((APRegistrationRsp)(results[0]));
        }
        
        /// <remarks/>
        public void APRegistrationAsync(APRegistrationReq APRegistrationReq) {
            this.APRegistrationAsync(APRegistrationReq, null);
        }
        
        /// <remarks/>
        public void APRegistrationAsync(APRegistrationReq APRegistrationReq, object userState) {
            if ((this.APRegistrationOperationCompleted == null)) {
                this.APRegistrationOperationCompleted = new System.Threading.SendOrPostCallback(this.OnAPRegistrationOperationCompleted);
            }
            this.InvokeAsync("APRegistration", new object[] {
                        APRegistrationReq}, this.APRegistrationOperationCompleted, userState);
        }
        
        private void OnAPRegistrationOperationCompleted(object arg) {
            if ((this.APRegistrationCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.APRegistrationCompleted(this, new APRegistrationCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/APStatusRep", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("APStatusRepRsp", Namespace="http://www.csapi.org/schema/ap")]
        public APStatusRepRsp APStatusRep([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ap")] APStatusRepReq APStatusRepReq) {
            object[] results = this.Invoke("APStatusRep", new object[] {
                        APStatusRepReq});
            return ((APStatusRepRsp)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginAPStatusRep(APStatusRepReq APStatusRepReq, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("APStatusRep", new object[] {
                        APStatusRepReq}, callback, asyncState);
        }
        
        /// <remarks/>
        public APStatusRepRsp EndAPStatusRep(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((APStatusRepRsp)(results[0]));
        }
        
        /// <remarks/>
        public void APStatusRepAsync(APStatusRepReq APStatusRepReq) {
            this.APStatusRepAsync(APStatusRepReq, null);
        }
        
        /// <remarks/>
        public void APStatusRepAsync(APStatusRepReq APStatusRepReq, object userState) {
            if ((this.APStatusRepOperationCompleted == null)) {
                this.APStatusRepOperationCompleted = new System.Threading.SendOrPostCallback(this.OnAPStatusRepOperationCompleted);
            }
            this.InvokeAsync("APStatusRep", new object[] {
                        APStatusRepReq}, this.APStatusRepOperationCompleted, userState);
        }
        
        private void OnAPStatusRepOperationCompleted(object arg) {
            if ((this.APStatusRepCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.APStatusRepCompleted(this, new APStatusRepCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/APLogOut", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("APLogOutRsp", Namespace="http://www.csapi.org/schema/ap")]
        public APLogOutRsp APLogOut([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ap")] APLogOutReq APLogOutReq) {
            object[] results = this.Invoke("APLogOut", new object[] {
                        APLogOutReq});
            return ((APLogOutRsp)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginAPLogOut(APLogOutReq APLogOutReq, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("APLogOut", new object[] {
                        APLogOutReq}, callback, asyncState);
        }
        
        /// <remarks/>
        public APLogOutRsp EndAPLogOut(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((APLogOutRsp)(results[0]));
        }
        
        /// <remarks/>
        public void APLogOutAsync(APLogOutReq APLogOutReq) {
            this.APLogOutAsync(APLogOutReq, null);
        }
        
        /// <remarks/>
        public void APLogOutAsync(APLogOutReq APLogOutReq, object userState) {
            if ((this.APLogOutOperationCompleted == null)) {
                this.APLogOutOperationCompleted = new System.Threading.SendOrPostCallback(this.OnAPLogOutOperationCompleted);
            }
            this.InvokeAsync("APLogOut", new object[] {
                        APLogOutReq}, this.APLogOutOperationCompleted, userState);
        }
        
        private void OnAPLogOutOperationCompleted(object arg) {
            if ((this.APLogOutCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.APLogOutCompleted(this, new APLogOutCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/PauseAP", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("PauseAPRsp", Namespace="http://www.csapi.org/schema/ap")]
        public PauseAPRsp PauseAP([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ap")] PauseAPReq PauseAPReq) {
            object[] results = this.Invoke("PauseAP", new object[] {
                        PauseAPReq});
            return ((PauseAPRsp)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginPauseAP(PauseAPReq PauseAPReq, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("PauseAP", new object[] {
                        PauseAPReq}, callback, asyncState);
        }
        
        /// <remarks/>
        public PauseAPRsp EndPauseAP(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((PauseAPRsp)(results[0]));
        }
        
        /// <remarks/>
        public void PauseAPAsync(PauseAPReq PauseAPReq) {
            this.PauseAPAsync(PauseAPReq, null);
        }
        
        /// <remarks/>
        public void PauseAPAsync(PauseAPReq PauseAPReq, object userState) {
            if ((this.PauseAPOperationCompleted == null)) {
                this.PauseAPOperationCompleted = new System.Threading.SendOrPostCallback(this.OnPauseAPOperationCompleted);
            }
            this.InvokeAsync("PauseAP", new object[] {
                        PauseAPReq}, this.PauseAPOperationCompleted, userState);
        }
        
        private void OnPauseAPOperationCompleted(object arg) {
            if ((this.PauseAPCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.PauseAPCompleted(this, new PauseAPCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/RecoveryAP", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("RecoveryAPRsp", Namespace="http://www.csapi.org/schema/ap")]
        public RecoveryAPRsp RecoveryAP([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ap")] RecoveryAPReq RecoveryAPReq) {
            object[] results = this.Invoke("RecoveryAP", new object[] {
                        RecoveryAPReq});
            return ((RecoveryAPRsp)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginRecoveryAP(RecoveryAPReq RecoveryAPReq, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("RecoveryAP", new object[] {
                        RecoveryAPReq}, callback, asyncState);
        }
        
        /// <remarks/>
        public RecoveryAPRsp EndRecoveryAP(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((RecoveryAPRsp)(results[0]));
        }
        
        /// <remarks/>
        public void RecoveryAPAsync(RecoveryAPReq RecoveryAPReq) {
            this.RecoveryAPAsync(RecoveryAPReq, null);
        }
        
        /// <remarks/>
        public void RecoveryAPAsync(RecoveryAPReq RecoveryAPReq, object userState) {
            if ((this.RecoveryAPOperationCompleted == null)) {
                this.RecoveryAPOperationCompleted = new System.Threading.SendOrPostCallback(this.OnRecoveryAPOperationCompleted);
            }
            this.InvokeAsync("RecoveryAP", new object[] {
                        RecoveryAPReq}, this.RecoveryAPOperationCompleted, userState);
        }
        
        private void OnRecoveryAPOperationCompleted(object arg) {
            if ((this.RecoveryAPCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.RecoveryAPCompleted(this, new RecoveryAPCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/APSvcAuthentic", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("APSvcAuthenticRsp", Namespace="http://www.csapi.org/schema/ap")]
        public APSvcAuthenticRsp APSvcAuthentic([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ap")] APSvcAuthenticReq APSvcAuthenticReq) {
            object[] results = this.Invoke("APSvcAuthentic", new object[] {
                        APSvcAuthenticReq});
            return ((APSvcAuthenticRsp)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginAPSvcAuthentic(APSvcAuthenticReq APSvcAuthenticReq, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("APSvcAuthentic", new object[] {
                        APSvcAuthenticReq}, callback, asyncState);
        }
        
        /// <remarks/>
        public APSvcAuthenticRsp EndAPSvcAuthentic(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((APSvcAuthenticRsp)(results[0]));
        }
        
        /// <remarks/>
        public void APSvcAuthenticAsync(APSvcAuthenticReq APSvcAuthenticReq) {
            this.APSvcAuthenticAsync(APSvcAuthenticReq, null);
        }
        
        /// <remarks/>
        public void APSvcAuthenticAsync(APSvcAuthenticReq APSvcAuthenticReq, object userState) {
            if ((this.APSvcAuthenticOperationCompleted == null)) {
                this.APSvcAuthenticOperationCompleted = new System.Threading.SendOrPostCallback(this.OnAPSvcAuthenticOperationCompleted);
            }
            this.InvokeAsync("APSvcAuthentic", new object[] {
                        APSvcAuthenticReq}, this.APSvcAuthenticOperationCompleted, userState);
        }
        
        private void OnAPSvcAuthenticOperationCompleted(object arg) {
            if ((this.APSvcAuthenticCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.APSvcAuthenticCompleted(this, new APSvcAuthenticCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/APSvcPerfCmd", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void APSvcPerfCmd([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ap")] APSvcPerfCmdReq APSvcPerfCmdReq) {
            this.Invoke("APSvcPerfCmd", new object[] {
                        APSvcPerfCmdReq});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginAPSvcPerfCmd(APSvcPerfCmdReq APSvcPerfCmdReq, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("APSvcPerfCmd", new object[] {
                        APSvcPerfCmdReq}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndAPSvcPerfCmd(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void APSvcPerfCmdAsync(APSvcPerfCmdReq APSvcPerfCmdReq) {
            this.APSvcPerfCmdAsync(APSvcPerfCmdReq, null);
        }
        
        /// <remarks/>
        public void APSvcPerfCmdAsync(APSvcPerfCmdReq APSvcPerfCmdReq, object userState) {
            if ((this.APSvcPerfCmdOperationCompleted == null)) {
                this.APSvcPerfCmdOperationCompleted = new System.Threading.SendOrPostCallback(this.OnAPSvcPerfCmdOperationCompleted);
            }
            this.InvokeAsync("APSvcPerfCmd", new object[] {
                        APSvcPerfCmdReq}, this.APSvcPerfCmdOperationCompleted, userState);
        }
        
        private void OnAPSvcPerfCmdOperationCompleted(object arg) {
            if ((this.APSvcPerfCmdCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.APSvcPerfCmdCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/APSvcAlarm", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("AlarmRsp", Namespace="http://www.csapi.org/schema/ap")]
        public AlarmRsp APSvcAlarm([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ap")] AlarmReq AlarmReq) {
            object[] results = this.Invoke("APSvcAlarm", new object[] {
                        AlarmReq});
            return ((AlarmRsp)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginAPSvcAlarm(AlarmReq AlarmReq, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("APSvcAlarm", new object[] {
                        AlarmReq}, callback, asyncState);
        }
        
        /// <remarks/>
        public AlarmRsp EndAPSvcAlarm(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((AlarmRsp)(results[0]));
        }
        
        /// <remarks/>
        public void APSvcAlarmAsync(AlarmReq AlarmReq) {
            this.APSvcAlarmAsync(AlarmReq, null);
        }
        
        /// <remarks/>
        public void APSvcAlarmAsync(AlarmReq AlarmReq, object userState) {
            if ((this.APSvcAlarmOperationCompleted == null)) {
                this.APSvcAlarmOperationCompleted = new System.Threading.SendOrPostCallback(this.OnAPSvcAlarmOperationCompleted);
            }
            this.InvokeAsync("APSvcAlarm", new object[] {
                        AlarmReq}, this.APSvcAlarmOperationCompleted, userState);
        }
        
        private void OnAPSvcAlarmOperationCompleted(object arg) {
            if ((this.APSvcAlarmCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.APSvcAlarmCompleted(this, new APSvcAlarmCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/APSvcPerfReport", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void APSvcPerfReport([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ap")] APSvcPerfReportReq APSvcPerfReportReq) {
            this.Invoke("APSvcPerfReport", new object[] {
                        APSvcPerfReportReq});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginAPSvcPerfReport(APSvcPerfReportReq APSvcPerfReportReq, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("APSvcPerfReport", new object[] {
                        APSvcPerfReportReq}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndAPSvcPerfReport(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void APSvcPerfReportAsync(APSvcPerfReportReq APSvcPerfReportReq) {
            this.APSvcPerfReportAsync(APSvcPerfReportReq, null);
        }
        
        /// <remarks/>
        public void APSvcPerfReportAsync(APSvcPerfReportReq APSvcPerfReportReq, object userState) {
            if ((this.APSvcPerfReportOperationCompleted == null)) {
                this.APSvcPerfReportOperationCompleted = new System.Threading.SendOrPostCallback(this.OnAPSvcPerfReportOperationCompleted);
            }
            this.InvokeAsync("APSvcPerfReport", new object[] {
                        APSvcPerfReportReq}, this.APSvcPerfReportOperationCompleted, userState);
        }
        
        private void OnAPSvcPerfReportOperationCompleted(object arg) {
            if ((this.APSvcPerfReportCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.APSvcPerfReportCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/startNotification", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void startNotification([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/notification")] startNotificationRequest startNotificationRequest) {
            this.Invoke("startNotification", new object[] {
                        startNotificationRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginstartNotification(startNotificationRequest startNotificationRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("startNotification", new object[] {
                        startNotificationRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndstartNotification(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void startNotificationAsync(startNotificationRequest startNotificationRequest) {
            this.startNotificationAsync(startNotificationRequest, null);
        }
        
        /// <remarks/>
        public void startNotificationAsync(startNotificationRequest startNotificationRequest, object userState) {
            if ((this.startNotificationOperationCompleted == null)) {
                this.startNotificationOperationCompleted = new System.Threading.SendOrPostCallback(this.OnstartNotificationOperationCompleted);
            }
            this.InvokeAsync("startNotification", new object[] {
                        startNotificationRequest}, this.startNotificationOperationCompleted, userState);
        }
        
        private void OnstartNotificationOperationCompleted(object arg) {
            if ((this.startNotificationCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.startNotificationCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/stopNotification", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void stopNotification([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/notification")] stopNotificationRequest stopNotificationRequest) {
            this.Invoke("stopNotification", new object[] {
                        stopNotificationRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginstopNotification(stopNotificationRequest stopNotificationRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("stopNotification", new object[] {
                        stopNotificationRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndstopNotification(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void stopNotificationAsync(stopNotificationRequest stopNotificationRequest) {
            this.stopNotificationAsync(stopNotificationRequest, null);
        }
        
        /// <remarks/>
        public void stopNotificationAsync(stopNotificationRequest stopNotificationRequest, object userState) {
            if ((this.stopNotificationOperationCompleted == null)) {
                this.stopNotificationOperationCompleted = new System.Threading.SendOrPostCallback(this.OnstopNotificationOperationCompleted);
            }
            this.InvokeAsync("stopNotification", new object[] {
                        stopNotificationRequest}, this.stopNotificationOperationCompleted, userState);
        }
        
        private void OnstopNotificationOperationCompleted(object arg) {
            if ((this.stopNotificationCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.stopNotificationCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/sendSms", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("sendSmsResponse", Namespace="http://www.csapi.org/schema/sms")]
        public sendSmsResponse sendSms([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/sms")] sendSmsRequest sendSmsRequest) {
            object[] results = this.Invoke("sendSms", new object[] {
                        sendSmsRequest});
            return ((sendSmsResponse)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginsendSms(sendSmsRequest sendSmsRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("sendSms", new object[] {
                        sendSmsRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public sendSmsResponse EndsendSms(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((sendSmsResponse)(results[0]));
        }
        
        /// <remarks/>
        public void sendSmsAsync(sendSmsRequest sendSmsRequest) {
            this.sendSmsAsync(sendSmsRequest, null);
        }
        
        /// <remarks/>
        public void sendSmsAsync(sendSmsRequest sendSmsRequest, object userState) {
            if ((this.sendSmsOperationCompleted == null)) {
                this.sendSmsOperationCompleted = new System.Threading.SendOrPostCallback(this.OnsendSmsOperationCompleted);
            }
            this.InvokeAsync("sendSms", new object[] {
                        sendSmsRequest}, this.sendSmsOperationCompleted, userState);
        }
        
        private void OnsendSmsOperationCompleted(object arg) {
            if ((this.sendSmsCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.sendSmsCompleted(this, new sendSmsCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/notifySmsDeliveryStatus", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void notifySmsDeliveryStatus([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/sms")] notifySmsDeliveryStatusRequest notifySmsDeliveryStatusRequest) {
            this.Invoke("notifySmsDeliveryStatus", new object[] {
                        notifySmsDeliveryStatusRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginnotifySmsDeliveryStatus(notifySmsDeliveryStatusRequest notifySmsDeliveryStatusRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("notifySmsDeliveryStatus", new object[] {
                        notifySmsDeliveryStatusRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndnotifySmsDeliveryStatus(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void notifySmsDeliveryStatusAsync(notifySmsDeliveryStatusRequest notifySmsDeliveryStatusRequest) {
            this.notifySmsDeliveryStatusAsync(notifySmsDeliveryStatusRequest, null);
        }
        
        /// <remarks/>
        public void notifySmsDeliveryStatusAsync(notifySmsDeliveryStatusRequest notifySmsDeliveryStatusRequest, object userState) {
            if ((this.notifySmsDeliveryStatusOperationCompleted == null)) {
                this.notifySmsDeliveryStatusOperationCompleted = new System.Threading.SendOrPostCallback(this.OnnotifySmsDeliveryStatusOperationCompleted);
            }
            this.InvokeAsync("notifySmsDeliveryStatus", new object[] {
                        notifySmsDeliveryStatusRequest}, this.notifySmsDeliveryStatusOperationCompleted, userState);
        }
        
        private void OnnotifySmsDeliveryStatusOperationCompleted(object arg) {
            if ((this.notifySmsDeliveryStatusCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.notifySmsDeliveryStatusCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/GetReceivedSms", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlArrayAttribute("GetReceivedSmsResponse", Namespace="http://www.csapi.org/schema/sms")]
        [return: System.Xml.Serialization.XmlArrayItemAttribute("ReceivedSms", Form = System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable = false)]
        public SMSMessage[] GetReceivedSms([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/sms")] GetReceivedSmsRequest GetReceivedSmsRequest) {
            object[] results = this.Invoke("GetReceivedSms", new object[] {
                        GetReceivedSmsRequest});
            return ((SMSMessage[])(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginGetReceivedSms(GetReceivedSmsRequest GetReceivedSmsRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("GetReceivedSms", new object[] {
                        GetReceivedSmsRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public SMSMessage[] EndGetReceivedSms(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((SMSMessage[])(results[0]));
        }
        
        /// <remarks/>
        public void GetReceivedSmsAsync(GetReceivedSmsRequest GetReceivedSmsRequest) {
            this.GetReceivedSmsAsync(GetReceivedSmsRequest, null);
        }
        
        /// <remarks/>
        public void GetReceivedSmsAsync(GetReceivedSmsRequest GetReceivedSmsRequest, object userState) {
            if ((this.GetReceivedSmsOperationCompleted == null)) {
                this.GetReceivedSmsOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetReceivedSmsOperationCompleted);
            }
            this.InvokeAsync("GetReceivedSms", new object[] {
                        GetReceivedSmsRequest}, this.GetReceivedSmsOperationCompleted, userState);
        }
        
        private void OnGetReceivedSmsOperationCompleted(object arg) {
            if ((this.GetReceivedSmsCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetReceivedSmsCompleted(this, new GetReceivedSmsCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/GetSmsDeliveryStatus", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlArrayAttribute("GetSmsDeliveryStatusResponse", Namespace="http://www.csapi.org/schema/sms")]
        [return: System.Xml.Serialization.XmlArrayItemAttribute("DeliveryStatus", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=false)]
        public DeliveryInformation[] GetSmsDeliveryStatus([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/sms")] GetSmsDeliveryStatusRequest GetSmsDeliveryStatusRequest) {
            object[] results = this.Invoke("GetSmsDeliveryStatus", new object[] {
                        GetSmsDeliveryStatusRequest});
            return ((DeliveryInformation[])(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginGetSmsDeliveryStatus(GetSmsDeliveryStatusRequest GetSmsDeliveryStatusRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("GetSmsDeliveryStatus", new object[] {
                        GetSmsDeliveryStatusRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public DeliveryInformation[] EndGetSmsDeliveryStatus(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((DeliveryInformation[])(results[0]));
        }
        
        /// <remarks/>
        public void GetSmsDeliveryStatusAsync(GetSmsDeliveryStatusRequest GetSmsDeliveryStatusRequest) {
            this.GetSmsDeliveryStatusAsync(GetSmsDeliveryStatusRequest, null);
        }
        
        /// <remarks/>
        public void GetSmsDeliveryStatusAsync(GetSmsDeliveryStatusRequest GetSmsDeliveryStatusRequest, object userState) {
            if ((this.GetSmsDeliveryStatusOperationCompleted == null)) {
                this.GetSmsDeliveryStatusOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetSmsDeliveryStatusOperationCompleted);
            }
            this.InvokeAsync("GetSmsDeliveryStatus", new object[] {
                        GetSmsDeliveryStatusRequest}, this.GetSmsDeliveryStatusOperationCompleted, userState);
        }
        
        private void OnGetSmsDeliveryStatusOperationCompleted(object arg) {
            if ((this.GetSmsDeliveryStatusCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetSmsDeliveryStatusCompleted(this, new GetSmsDeliveryStatusCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/notifySmsReception", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void notifySmsReception([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/sms")] notifySmsReceptionRequest notifySmsReceptionRequest) {
            this.Invoke("notifySmsReception", new object[] {
                        notifySmsReceptionRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginnotifySmsReception(notifySmsReceptionRequest notifySmsReceptionRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("notifySmsReception", new object[] {
                        notifySmsReceptionRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndnotifySmsReception(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void notifySmsReceptionAsync(notifySmsReceptionRequest notifySmsReceptionRequest) {
            this.notifySmsReceptionAsync(notifySmsReceptionRequest, null);
        }
        
        /// <remarks/>
        public void notifySmsReceptionAsync(notifySmsReceptionRequest notifySmsReceptionRequest, object userState) {
            if ((this.notifySmsReceptionOperationCompleted == null)) {
                this.notifySmsReceptionOperationCompleted = new System.Threading.SendOrPostCallback(this.OnnotifySmsReceptionOperationCompleted);
            }
            this.InvokeAsync("notifySmsReception", new object[] {
                        notifySmsReceptionRequest}, this.notifySmsReceptionOperationCompleted, userState);
        }
        
        private void OnnotifySmsReceptionOperationCompleted(object arg) {
            if ((this.notifySmsReceptionCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.notifySmsReceptionCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/sendMessage", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("sendMessageResponse", Namespace="http://www.csapi.org/schema/mms")]
        public sendMessageResponse sendMessage([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/mms")] sendMessageRequest sendMessageRequest) {
            object[] results = this.Invoke("sendMessage", new object[] {
                        sendMessageRequest});
            return ((sendMessageResponse)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginsendMessage(sendMessageRequest sendMessageRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("sendMessage", new object[] {
                        sendMessageRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public sendMessageResponse EndsendMessage(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((sendMessageResponse)(results[0]));
        }
        
        /// <remarks/>
        public void sendMessageAsync(sendMessageRequest sendMessageRequest) {
            this.sendMessageAsync(sendMessageRequest, null);
        }
        
        /// <remarks/>
        public void sendMessageAsync(sendMessageRequest sendMessageRequest, object userState) {
            if ((this.sendMessageOperationCompleted == null)) {
                this.sendMessageOperationCompleted = new System.Threading.SendOrPostCallback(this.OnsendMessageOperationCompleted);
            }
            this.InvokeAsync("sendMessage", new object[] {
                        sendMessageRequest}, this.sendMessageOperationCompleted, userState);
        }
        
        private void OnsendMessageOperationCompleted(object arg) {
            if ((this.sendMessageCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.sendMessageCompleted(this, new sendMessageCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/getMessageDeliveryStatus", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlArrayAttribute("getMessageDeliveryStatusResponse", Namespace="http://www.csapi.org/schema/mms")]
        [return: System.Xml.Serialization.XmlArrayItemAttribute("item", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=false)]
        public DeliveryInformation1[] getMessageDeliveryStatus([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/mms")] getMessageDeliveryStatusRequest getMessageDeliveryStatusRequest) {
            object[] results = this.Invoke("getMessageDeliveryStatus", new object[] {
                        getMessageDeliveryStatusRequest});
            return ((DeliveryInformation1[])(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BegingetMessageDeliveryStatus(getMessageDeliveryStatusRequest getMessageDeliveryStatusRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("getMessageDeliveryStatus", new object[] {
                        getMessageDeliveryStatusRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public DeliveryInformation1[] EndgetMessageDeliveryStatus(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((DeliveryInformation1[])(results[0]));
        }
        
        /// <remarks/>
        public void getMessageDeliveryStatusAsync(getMessageDeliveryStatusRequest getMessageDeliveryStatusRequest) {
            this.getMessageDeliveryStatusAsync(getMessageDeliveryStatusRequest, null);
        }
        
        /// <remarks/>
        public void getMessageDeliveryStatusAsync(getMessageDeliveryStatusRequest getMessageDeliveryStatusRequest, object userState) {
            if ((this.getMessageDeliveryStatusOperationCompleted == null)) {
                this.getMessageDeliveryStatusOperationCompleted = new System.Threading.SendOrPostCallback(this.OngetMessageDeliveryStatusOperationCompleted);
            }
            this.InvokeAsync("getMessageDeliveryStatus", new object[] {
                        getMessageDeliveryStatusRequest}, this.getMessageDeliveryStatusOperationCompleted, userState);
        }
        
        private void OngetMessageDeliveryStatusOperationCompleted(object arg) {
            if ((this.getMessageDeliveryStatusCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.getMessageDeliveryStatusCompleted(this, new getMessageDeliveryStatusCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/getReceivedMessages", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlArrayAttribute("getReceivedMessagesResponse", Namespace="http://www.csapi.org/schema/mms")]
        [return: System.Xml.Serialization.XmlArrayItemAttribute("item", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=false)]
        public MessageReference[] getReceivedMessages([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/mms")] getReceivedMessagesRequest getReceivedMessagesRequest) {
            object[] results = this.Invoke("getReceivedMessages", new object[] {
                        getReceivedMessagesRequest});
            return ((MessageReference[])(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BegingetReceivedMessages(getReceivedMessagesRequest getReceivedMessagesRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("getReceivedMessages", new object[] {
                        getReceivedMessagesRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public MessageReference[] EndgetReceivedMessages(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((MessageReference[])(results[0]));
        }
        
        /// <remarks/>
        public void getReceivedMessagesAsync(getReceivedMessagesRequest getReceivedMessagesRequest) {
            this.getReceivedMessagesAsync(getReceivedMessagesRequest, null);
        }
        
        /// <remarks/>
        public void getReceivedMessagesAsync(getReceivedMessagesRequest getReceivedMessagesRequest, object userState) {
            if ((this.getReceivedMessagesOperationCompleted == null)) {
                this.getReceivedMessagesOperationCompleted = new System.Threading.SendOrPostCallback(this.OngetReceivedMessagesOperationCompleted);
            }
            this.InvokeAsync("getReceivedMessages", new object[] {
                        getReceivedMessagesRequest}, this.getReceivedMessagesOperationCompleted, userState);
        }
        
        private void OngetReceivedMessagesOperationCompleted(object arg) {
            if ((this.getReceivedMessagesCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.getReceivedMessagesCompleted(this, new getReceivedMessagesCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/getMessage", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("getMessageResponse", Namespace="http://www.csapi.org/schema/mms")]
        public getMessageResponse getMessage([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/mms")] getMessageRequest getMessageRequest) {
            object[] results = this.Invoke("getMessage", new object[] {
                        getMessageRequest});
            return ((getMessageResponse)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BegingetMessage(getMessageRequest getMessageRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("getMessage", new object[] {
                        getMessageRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public getMessageResponse EndgetMessage(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((getMessageResponse)(results[0]));
        }
        
        /// <remarks/>
        public void getMessageAsync(getMessageRequest getMessageRequest) {
            this.getMessageAsync(getMessageRequest, null);
        }
        
        /// <remarks/>
        public void getMessageAsync(getMessageRequest getMessageRequest, object userState) {
            if ((this.getMessageOperationCompleted == null)) {
                this.getMessageOperationCompleted = new System.Threading.SendOrPostCallback(this.OngetMessageOperationCompleted);
            }
            this.InvokeAsync("getMessage", new object[] {
                        getMessageRequest}, this.getMessageOperationCompleted, userState);
        }
        
        private void OngetMessageOperationCompleted(object arg) {
            if ((this.getMessageCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.getMessageCompleted(this, new getMessageCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/notifyMessageReception", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void notifyMessageReception([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/mms")] notifyMessageReceptionRequest notifyMessageReceptionRequest) {
            this.Invoke("notifyMessageReception", new object[] {
                        notifyMessageReceptionRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginnotifyMessageReception(notifyMessageReceptionRequest notifyMessageReceptionRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("notifyMessageReception", new object[] {
                        notifyMessageReceptionRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndnotifyMessageReception(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void notifyMessageReceptionAsync(notifyMessageReceptionRequest notifyMessageReceptionRequest) {
            this.notifyMessageReceptionAsync(notifyMessageReceptionRequest, null);
        }
        
        /// <remarks/>
        public void notifyMessageReceptionAsync(notifyMessageReceptionRequest notifyMessageReceptionRequest, object userState) {
            if ((this.notifyMessageReceptionOperationCompleted == null)) {
                this.notifyMessageReceptionOperationCompleted = new System.Threading.SendOrPostCallback(this.OnnotifyMessageReceptionOperationCompleted);
            }
            this.InvokeAsync("notifyMessageReception", new object[] {
                        notifyMessageReceptionRequest}, this.notifyMessageReceptionOperationCompleted, userState);
        }
        
        private void OnnotifyMessageReceptionOperationCompleted(object arg) {
            if ((this.notifyMessageReceptionCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.notifyMessageReceptionCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/notifyMessageDeliveryReceipt", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void notifyMessageDeliveryReceipt([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/mms")] notifyMessageDeliveryReceiptRequest notifyMessageDeliveryReceiptRequest) {
            this.Invoke("notifyMessageDeliveryReceipt", new object[] {
                        notifyMessageDeliveryReceiptRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginnotifyMessageDeliveryReceipt(notifyMessageDeliveryReceiptRequest notifyMessageDeliveryReceiptRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("notifyMessageDeliveryReceipt", new object[] {
                        notifyMessageDeliveryReceiptRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndnotifyMessageDeliveryReceipt(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void notifyMessageDeliveryReceiptAsync(notifyMessageDeliveryReceiptRequest notifyMessageDeliveryReceiptRequest) {
            this.notifyMessageDeliveryReceiptAsync(notifyMessageDeliveryReceiptRequest, null);
        }
        
        /// <remarks/>
        public void notifyMessageDeliveryReceiptAsync(notifyMessageDeliveryReceiptRequest notifyMessageDeliveryReceiptRequest, object userState) {
            if ((this.notifyMessageDeliveryReceiptOperationCompleted == null)) {
                this.notifyMessageDeliveryReceiptOperationCompleted = new System.Threading.SendOrPostCallback(this.OnnotifyMessageDeliveryReceiptOperationCompleted);
            }
            this.InvokeAsync("notifyMessageDeliveryReceipt", new object[] {
                        notifyMessageDeliveryReceiptRequest}, this.notifyMessageDeliveryReceiptOperationCompleted, userState);
        }
        
        private void OnnotifyMessageDeliveryReceiptOperationCompleted(object arg) {
            if ((this.notifyMessageDeliveryReceiptCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.notifyMessageDeliveryReceiptCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/getLocation", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("getLocationResponse", Namespace="http://www.csapi.org/schema/location")]
        public getLocationResponse getLocation([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/location")] getLocationRequest getLocationRequest) {
            object[] results = this.Invoke("getLocation", new object[] {
                        getLocationRequest});
            return ((getLocationResponse)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BegingetLocation(getLocationRequest getLocationRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("getLocation", new object[] {
                        getLocationRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public getLocationResponse EndgetLocation(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((getLocationResponse)(results[0]));
        }
        
        /// <remarks/>
        public void getLocationAsync(getLocationRequest getLocationRequest) {
            this.getLocationAsync(getLocationRequest, null);
        }
        
        /// <remarks/>
        public void getLocationAsync(getLocationRequest getLocationRequest, object userState) {
            if ((this.getLocationOperationCompleted == null)) {
                this.getLocationOperationCompleted = new System.Threading.SendOrPostCallback(this.OngetLocationOperationCompleted);
            }
            this.InvokeAsync("getLocation", new object[] {
                        getLocationRequest}, this.getLocationOperationCompleted, userState);
        }
        
        private void OngetLocationOperationCompleted(object arg) {
            if ((this.getLocationCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.getLocationCompleted(this, new getLocationCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/getLocationForGroup", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlArrayAttribute("getLocationForGroupResponse", Namespace="http://www.csapi.org/schema/location")]
        [return: System.Xml.Serialization.XmlArrayItemAttribute("item", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=false)]
        public LocationData[] getLocationForGroup([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/location")] getLocationForGroupRequest getLocationForGroupRequest) {
            object[] results = this.Invoke("getLocationForGroup", new object[] {
                        getLocationForGroupRequest});
            return ((LocationData[])(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BegingetLocationForGroup(getLocationForGroupRequest getLocationForGroupRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("getLocationForGroup", new object[] {
                        getLocationForGroupRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public LocationData[] EndgetLocationForGroup(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((LocationData[])(results[0]));
        }
        
        /// <remarks/>
        public void getLocationForGroupAsync(getLocationForGroupRequest getLocationForGroupRequest) {
            this.getLocationForGroupAsync(getLocationForGroupRequest, null);
        }
        
        /// <remarks/>
        public void getLocationForGroupAsync(getLocationForGroupRequest getLocationForGroupRequest, object userState) {
            if ((this.getLocationForGroupOperationCompleted == null)) {
                this.getLocationForGroupOperationCompleted = new System.Threading.SendOrPostCallback(this.OngetLocationForGroupOperationCompleted);
            }
            this.InvokeAsync("getLocationForGroup", new object[] {
                        getLocationForGroupRequest}, this.getLocationForGroupOperationCompleted, userState);
        }
        
        private void OngetLocationForGroupOperationCompleted(object arg) {
            if ((this.getLocationForGroupCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.getLocationForGroupCompleted(this, new getLocationForGroupCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/startPeriodicNotification", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("startPeriodicNotificationResponse", Namespace="http://www.csapi.org/schema/location")]
        public startPeriodicNotificationResponse startPeriodicNotification([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/location")] startPeriodicNotificationRequest startPeriodicNotificationRequest) {
            object[] results = this.Invoke("startPeriodicNotification", new object[] {
                        startPeriodicNotificationRequest});
            return ((startPeriodicNotificationResponse)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginstartPeriodicNotification(startPeriodicNotificationRequest startPeriodicNotificationRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("startPeriodicNotification", new object[] {
                        startPeriodicNotificationRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public startPeriodicNotificationResponse EndstartPeriodicNotification(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((startPeriodicNotificationResponse)(results[0]));
        }
        
        /// <remarks/>
        public void startPeriodicNotificationAsync(startPeriodicNotificationRequest startPeriodicNotificationRequest) {
            this.startPeriodicNotificationAsync(startPeriodicNotificationRequest, null);
        }
        
        /// <remarks/>
        public void startPeriodicNotificationAsync(startPeriodicNotificationRequest startPeriodicNotificationRequest, object userState) {
            if ((this.startPeriodicNotificationOperationCompleted == null)) {
                this.startPeriodicNotificationOperationCompleted = new System.Threading.SendOrPostCallback(this.OnstartPeriodicNotificationOperationCompleted);
            }
            this.InvokeAsync("startPeriodicNotification", new object[] {
                        startPeriodicNotificationRequest}, this.startPeriodicNotificationOperationCompleted, userState);
        }
        
        private void OnstartPeriodicNotificationOperationCompleted(object arg) {
            if ((this.startPeriodicNotificationCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.startPeriodicNotificationCompleted(this, new startPeriodicNotificationCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/EndNotification", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void EndNotification([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/location")] EndNotificationRequest EndNotificationRequest) {
            this.Invoke("EndNotification", new object[] {
                        EndNotificationRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginEndNotification(EndNotificationRequest EndNotificationRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("EndNotification", new object[] {
                        EndNotificationRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndEndNotification(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void EndNotificationAsync(EndNotificationRequest EndNotificationRequest) {
            this.EndNotificationAsync(EndNotificationRequest, null);
        }
        
        /// <remarks/>
        public void EndNotificationAsync(EndNotificationRequest EndNotificationRequest, object userState) {
            if ((this.EndNotificationOperationCompleted == null)) {
                this.EndNotificationOperationCompleted = new System.Threading.SendOrPostCallback(this.OnEndNotificationOperationCompleted);
            }
            this.InvokeAsync("EndNotification", new object[] {
                        EndNotificationRequest}, this.EndNotificationOperationCompleted, userState);
        }
        
        private void OnEndNotificationOperationCompleted(object arg) {
            if ((this.EndNotificationCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.EndNotificationCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/LocationNotification", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void LocationNotification([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/location")] LocationNotificationRequest LocationNotificationRequest) {
            this.Invoke("LocationNotification", new object[] {
                        LocationNotificationRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginLocationNotification(LocationNotificationRequest LocationNotificationRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("LocationNotification", new object[] {
                        LocationNotificationRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndLocationNotification(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void LocationNotificationAsync(LocationNotificationRequest LocationNotificationRequest) {
            this.LocationNotificationAsync(LocationNotificationRequest, null);
        }
        
        /// <remarks/>
        public void LocationNotificationAsync(LocationNotificationRequest LocationNotificationRequest, object userState) {
            if ((this.LocationNotificationOperationCompleted == null)) {
                this.LocationNotificationOperationCompleted = new System.Threading.SendOrPostCallback(this.OnLocationNotificationOperationCompleted);
            }
            this.InvokeAsync("LocationNotification", new object[] {
                        LocationNotificationRequest}, this.LocationNotificationOperationCompleted, userState);
        }
        
        private void OnLocationNotificationOperationCompleted(object arg) {
            if ((this.LocationNotificationCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.LocationNotificationCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/LocationError", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void LocationError([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/location")] LocationErrorRequest LocationErrorRequest) {
            this.Invoke("LocationError", new object[] {
                        LocationErrorRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginLocationError(LocationErrorRequest LocationErrorRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("LocationError", new object[] {
                        LocationErrorRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndLocationError(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void LocationErrorAsync(LocationErrorRequest LocationErrorRequest) {
            this.LocationErrorAsync(LocationErrorRequest, null);
        }
        
        /// <remarks/>
        public void LocationErrorAsync(LocationErrorRequest LocationErrorRequest, object userState) {
            if ((this.LocationErrorOperationCompleted == null)) {
                this.LocationErrorOperationCompleted = new System.Threading.SendOrPostCallback(this.OnLocationErrorOperationCompleted);
            }
            this.InvokeAsync("LocationError", new object[] {
                        LocationErrorRequest}, this.LocationErrorOperationCompleted, userState);
        }
        
        private void OnLocationErrorOperationCompleted(object arg) {
            if ((this.LocationErrorCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.LocationErrorCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/LocationEnd", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void LocationEnd([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/location")] LocationEndRequest LocationEndRequest) {
            this.Invoke("LocationEnd", new object[] {
                        LocationEndRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginLocationEnd(LocationEndRequest LocationEndRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("LocationEnd", new object[] {
                        LocationEndRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndLocationEnd(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void LocationEndAsync(LocationEndRequest LocationEndRequest) {
            this.LocationEndAsync(LocationEndRequest, null);
        }
        
        /// <remarks/>
        public void LocationEndAsync(LocationEndRequest LocationEndRequest, object userState) {
            if ((this.LocationEndOperationCompleted == null)) {
                this.LocationEndOperationCompleted = new System.Threading.SendOrPostCallback(this.OnLocationEndOperationCompleted);
            }
            this.InvokeAsync("LocationEnd", new object[] {
                        LocationEndRequest}, this.LocationEndOperationCompleted, userState);
        }
        
        private void OnLocationEndOperationCompleted(object arg) {
            if ((this.LocationEndCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.LocationEndCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/sendPush", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("sendPushResponse", Namespace="http://www.csapi.org/schema/wap")]
        public sendPushResponse sendPush([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/wap")] sendPushRequest sendPushRequest) {
            object[] results = this.Invoke("sendPush", new object[] {
                        sendPushRequest});
            return ((sendPushResponse)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginsendPush(sendPushRequest sendPushRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("sendPush", new object[] {
                        sendPushRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public sendPushResponse EndsendPush(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((sendPushResponse)(results[0]));
        }
        
        /// <remarks/>
        public void sendPushAsync(sendPushRequest sendPushRequest) {
            this.sendPushAsync(sendPushRequest, null);
        }
        
        /// <remarks/>
        public void sendPushAsync(sendPushRequest sendPushRequest, object userState) {
            if ((this.sendPushOperationCompleted == null)) {
                this.sendPushOperationCompleted = new System.Threading.SendOrPostCallback(this.OnsendPushOperationCompleted);
            }
            this.InvokeAsync("sendPush", new object[] {
                        sendPushRequest}, this.sendPushOperationCompleted, userState);
        }
        
        private void OnsendPushOperationCompleted(object arg) {
            if ((this.sendPushCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.sendPushCompleted(this, new sendPushCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/getPushDeliveryStatus", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlArrayAttribute("getPushDeliveryStatusResponse", Namespace="http://www.csapi.org/schema/wap")]
        [return: System.Xml.Serialization.XmlArrayItemAttribute("item", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=false)]
        public DeliveryInformation2[] getPushDeliveryStatus([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/wap")] getPushDeliveryStatusRequest getPushDeliveryStatusRequest) {
            object[] results = this.Invoke("getPushDeliveryStatus", new object[] {
                        getPushDeliveryStatusRequest});
            return ((DeliveryInformation2[])(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BegingetPushDeliveryStatus(getPushDeliveryStatusRequest getPushDeliveryStatusRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("getPushDeliveryStatus", new object[] {
                        getPushDeliveryStatusRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public DeliveryInformation2[] EndgetPushDeliveryStatus(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((DeliveryInformation2[])(results[0]));
        }
        
        /// <remarks/>
        public void getPushDeliveryStatusAsync(getPushDeliveryStatusRequest getPushDeliveryStatusRequest) {
            this.getPushDeliveryStatusAsync(getPushDeliveryStatusRequest, null);
        }
        
        /// <remarks/>
        public void getPushDeliveryStatusAsync(getPushDeliveryStatusRequest getPushDeliveryStatusRequest, object userState) {
            if ((this.getPushDeliveryStatusOperationCompleted == null)) {
                this.getPushDeliveryStatusOperationCompleted = new System.Threading.SendOrPostCallback(this.OngetPushDeliveryStatusOperationCompleted);
            }
            this.InvokeAsync("getPushDeliveryStatus", new object[] {
                        getPushDeliveryStatusRequest}, this.getPushDeliveryStatusOperationCompleted, userState);
        }
        
        private void OngetPushDeliveryStatusOperationCompleted(object arg) {
            if ((this.getPushDeliveryStatusCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.getPushDeliveryStatusCompleted(this, new getPushDeliveryStatusCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/notifyPushDeliveryReceipt", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void notifyPushDeliveryReceipt([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/wap")] notifyPushDeliveryReceiptRequest notifyPushDeliveryReceiptRequest) {
            this.Invoke("notifyPushDeliveryReceipt", new object[] {
                        notifyPushDeliveryReceiptRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginnotifyPushDeliveryReceipt(notifyPushDeliveryReceiptRequest notifyPushDeliveryReceiptRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("notifyPushDeliveryReceipt", new object[] {
                        notifyPushDeliveryReceiptRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndnotifyPushDeliveryReceipt(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void notifyPushDeliveryReceiptAsync(notifyPushDeliveryReceiptRequest notifyPushDeliveryReceiptRequest) {
            this.notifyPushDeliveryReceiptAsync(notifyPushDeliveryReceiptRequest, null);
        }
        
        /// <remarks/>
        public void notifyPushDeliveryReceiptAsync(notifyPushDeliveryReceiptRequest notifyPushDeliveryReceiptRequest, object userState) {
            if ((this.notifyPushDeliveryReceiptOperationCompleted == null)) {
                this.notifyPushDeliveryReceiptOperationCompleted = new System.Threading.SendOrPostCallback(this.OnnotifyPushDeliveryReceiptOperationCompleted);
            }
            this.InvokeAsync("notifyPushDeliveryReceipt", new object[] {
                        notifyPushDeliveryReceiptRequest}, this.notifyPushDeliveryReceiptOperationCompleted, userState);
        }
        
        private void OnnotifyPushDeliveryReceiptOperationCompleted(object arg) {
            if ((this.notifyPushDeliveryReceiptCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.notifyPushDeliveryReceiptCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/makeUssd", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("makeUssdResponse", Namespace="http://www.csapi.org/schema/ussd")]
        public makeUssdResponse makeUssd([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ussd")] makeUssdRequest makeUssdRequest) {
            object[] results = this.Invoke("makeUssd", new object[] {
                        makeUssdRequest});
            return ((makeUssdResponse)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginmakeUssd(makeUssdRequest makeUssdRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("makeUssd", new object[] {
                        makeUssdRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public makeUssdResponse EndmakeUssd(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((makeUssdResponse)(results[0]));
        }
        
        /// <remarks/>
        public void makeUssdAsync(makeUssdRequest makeUssdRequest) {
            this.makeUssdAsync(makeUssdRequest, null);
        }
        
        /// <remarks/>
        public void makeUssdAsync(makeUssdRequest makeUssdRequest, object userState) {
            if ((this.makeUssdOperationCompleted == null)) {
                this.makeUssdOperationCompleted = new System.Threading.SendOrPostCallback(this.OnmakeUssdOperationCompleted);
            }
            this.InvokeAsync("makeUssd", new object[] {
                        makeUssdRequest}, this.makeUssdOperationCompleted, userState);
        }
        
        private void OnmakeUssdOperationCompleted(object arg) {
            if ((this.makeUssdCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.makeUssdCompleted(this, new makeUssdCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/handleUssd", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("handleUssdResponse", Namespace="http://www.csapi.org/schema/ussd")]
        public handleUssdResponse handleUssd([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ussd")] handleUssdRequest handleUssdRequest) {
            object[] results = this.Invoke("handleUssd", new object[] {
                        handleUssdRequest});
            return ((handleUssdResponse)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginhandleUssd(handleUssdRequest handleUssdRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("handleUssd", new object[] {
                        handleUssdRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public handleUssdResponse EndhandleUssd(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((handleUssdResponse)(results[0]));
        }
        
        /// <remarks/>
        public void handleUssdAsync(handleUssdRequest handleUssdRequest) {
            this.handleUssdAsync(handleUssdRequest, null);
        }
        
        /// <remarks/>
        public void handleUssdAsync(handleUssdRequest handleUssdRequest, object userState) {
            if ((this.handleUssdOperationCompleted == null)) {
                this.handleUssdOperationCompleted = new System.Threading.SendOrPostCallback(this.OnhandleUssdOperationCompleted);
            }
            this.InvokeAsync("handleUssd", new object[] {
                        handleUssdRequest}, this.handleUssdOperationCompleted, userState);
        }
        
        private void OnhandleUssdOperationCompleted(object arg) {
            if ((this.handleUssdCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.handleUssdCompleted(this, new handleUssdCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/ussdContinue", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("ussdContinueResponse", Namespace="http://www.csapi.org/schema/ussd")]
        public ussdContinueResponse ussdContinue([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ussd")] ussdContinueRequest ussdContinueRequest) {
            object[] results = this.Invoke("ussdContinue", new object[] {
                        ussdContinueRequest});
            return ((ussdContinueResponse)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginussdContinue(ussdContinueRequest ussdContinueRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("ussdContinue", new object[] {
                        ussdContinueRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public ussdContinueResponse EndussdContinue(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((ussdContinueResponse)(results[0]));
        }
        
        /// <remarks/>
        public void ussdContinueAsync(ussdContinueRequest ussdContinueRequest) {
            this.ussdContinueAsync(ussdContinueRequest, null);
        }
        
        /// <remarks/>
        public void ussdContinueAsync(ussdContinueRequest ussdContinueRequest, object userState) {
            if ((this.ussdContinueOperationCompleted == null)) {
                this.ussdContinueOperationCompleted = new System.Threading.SendOrPostCallback(this.OnussdContinueOperationCompleted);
            }
            this.InvokeAsync("ussdContinue", new object[] {
                        ussdContinueRequest}, this.ussdContinueOperationCompleted, userState);
        }
        
        private void OnussdContinueOperationCompleted(object arg) {
            if ((this.ussdContinueCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.ussdContinueCompleted(this, new ussdContinueCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/endUssd", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void endUssd([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ussd")] endUssdRequest endUssdRequest) {
            this.Invoke("endUssd", new object[] {
                        endUssdRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginendUssd(endUssdRequest endUssdRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("endUssd", new object[] {
                        endUssdRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndendUssd(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void endUssdAsync(endUssdRequest endUssdRequest) {
            this.endUssdAsync(endUssdRequest, null);
        }
        
        /// <remarks/>
        public void endUssdAsync(endUssdRequest endUssdRequest, object userState) {
            if ((this.endUssdOperationCompleted == null)) {
                this.endUssdOperationCompleted = new System.Threading.SendOrPostCallback(this.OnendUssdOperationCompleted);
            }
            this.InvokeAsync("endUssd", new object[] {
                        endUssdRequest}, this.endUssdOperationCompleted, userState);
        }
        
        private void OnendUssdOperationCompleted(object arg) {
            if ((this.endUssdCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.endUssdCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.csapi.org/service/notifyUssdEnd", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        public void notifyUssdEnd([System.Xml.Serialization.XmlElementAttribute(Namespace="http://www.csapi.org/schema/ussd")] notifyUssdEndRequest notifyUssdEndRequest) {
            this.Invoke("notifyUssdEnd", new object[] {
                        notifyUssdEndRequest});
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginnotifyUssdEnd(notifyUssdEndRequest notifyUssdEndRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("notifyUssdEnd", new object[] {
                        notifyUssdEndRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public void EndnotifyUssdEnd(System.IAsyncResult asyncResult) {
            this.EndInvoke(asyncResult);
        }
        
        /// <remarks/>
        public void notifyUssdEndAsync(notifyUssdEndRequest notifyUssdEndRequest) {
            this.notifyUssdEndAsync(notifyUssdEndRequest, null);
        }
        
        /// <remarks/>
        public void notifyUssdEndAsync(notifyUssdEndRequest notifyUssdEndRequest, object userState) {
            if ((this.notifyUssdEndOperationCompleted == null)) {
                this.notifyUssdEndOperationCompleted = new System.Threading.SendOrPostCallback(this.OnnotifyUssdEndOperationCompleted);
            }
            this.InvokeAsync("notifyUssdEnd", new object[] {
                        notifyUssdEndRequest}, this.notifyUssdEndOperationCompleted, userState);
        }
        
        private void OnnotifyUssdEndOperationCompleted(object arg) {
            if ((this.notifyUssdEndCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.notifyUssdEndCompleted(this, new System.ComponentModel.AsyncCompletedEventArgs(invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        public new void CancelAsync(object userState) {
            base.CancelAsync(userState);
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class APRegistrationReq {
        
        private string apidField;
        
        private int aPPidField;
        
        private string hostIPField;
        
        private MessageNotificationType[] messageNotificationField;
        
        private string aPWSURIField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string Apid {
            get {
                return this.apidField;
            }
            set {
                this.apidField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int APPid {
            get {
                return this.aPPidField;
            }
            set {
                this.aPPidField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string HostIP {
            get {
                return this.hostIPField;
            }
            set {
                this.hostIPField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("MessageNotification", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public MessageNotificationType[] MessageNotification {
            get {
                return this.messageNotificationField;
            }
            set {
                this.messageNotificationField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI")]
        public string APWSURI {
            get {
                return this.aPWSURIField;
            }
            set {
                this.aPWSURIField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/common/v2_0")]
    public partial class MessageNotificationType {
        
        private CMAbility cMAbilityField;
        
        private string[] wSURIField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public CMAbility CMAbility {
            get {
                return this.cMAbilityField;
            }
            set {
                this.cMAbilityField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("WSURI", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI")]
        public string[] WSURI {
            get {
                return this.wSURIField;
            }
            set {
                this.wSURIField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/common/v2_0")]
    public enum CMAbility {
        
        /// <remarks/>
        SMSAbility,
        
        /// <remarks/>
        MMSAbility,
        
        /// <remarks/>
        WAPAbility,
        
        /// <remarks/>
        USSDAbility,
        
        /// <remarks/>
        LBSAbility,
        
        /// <remarks/>
        GPRSAbility,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/ussd")]
    public partial class UssdArray {
        
        private string ussdMessageField;
        
        private bool ussdReturnRequestField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ussdMessage {
            get {
                return this.ussdMessageField;
            }
            set {
                this.ussdMessageField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public bool ussdReturnRequest {
            get {
                return this.ussdReturnRequestField;
            }
            set {
                this.ussdReturnRequestField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(TypeName="DeliveryInformation", Namespace="http://www.csapi.org/schema/wap")]
    public partial class DeliveryInformation2 {
        
        private string addressField;
        
        private DeliveryStatus2 statusField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI")]
        public string address {
            get {
                return this.addressField;
            }
            set {
                this.addressField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public DeliveryStatus2 status {
            get {
                return this.statusField;
            }
            set {
                this.statusField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(TypeName="DeliveryStatus", Namespace="http://www.csapi.org/schema/wap")]
    public enum DeliveryStatus2 {
        
        /// <remarks/>
        DeliveredToNetwork,
        
        /// <remarks/>
        DeliveryUncertain,
        
        /// <remarks/>
        DeliveryImpossible,
        
        /// <remarks/>
        MessageWaiting,
        
        /// <remarks/>
        DeliveredToTerminal,
        
        /// <remarks/>
        DeliveryNotificationNotSupported,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/common/v2_0")]
    public partial class ServiceError {
        
        private string messageIdField;
        
        private string textField;
        
        private string[] variablesField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string MessageId {
            get {
                return this.messageIdField;
            }
            set {
                this.messageIdField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string Text {
            get {
                return this.textField;
            }
            set {
                this.textField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Variables", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string[] Variables {
            get {
                return this.variablesField;
            }
            set {
                this.variablesField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/location")]
    public partial class LocationData {
        
        private string addressField;
        
        private System.Nullable<RetrievalStatus> reportStatusField;
        
        private LocationInfo currentLocationField;
        
        private ServiceError errorInformationField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string Address {
            get {
                return this.addressField;
            }
            set {
                this.addressField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<RetrievalStatus> ReportStatus {
            get {
                return this.reportStatusField;
            }
            set {
                this.reportStatusField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public LocationInfo CurrentLocation {
            get {
                return this.currentLocationField;
            }
            set {
                this.currentLocationField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public ServiceError ErrorInformation {
            get {
                return this.errorInformationField;
            }
            set {
                this.errorInformationField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/location")]
    public enum RetrievalStatus {
        
        /// <remarks/>
        Retrieved,
        
        /// <remarks/>
        NotRetrieved,
        
        /// <remarks/>
        Error,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/location")]
    public partial class LocationInfo {
        
        private float latitudeField;
        
        private float longitudeField;
        
        private float altitudeField;
        
        private int accuracyField;
        
        private System.Nullable<System.DateTime> timestampField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public float Latitude {
            get {
                return this.latitudeField;
            }
            set {
                this.latitudeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public float Longitude {
            get {
                return this.longitudeField;
            }
            set {
                this.longitudeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public float Altitude {
            get {
                return this.altitudeField;
            }
            set {
                this.altitudeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int Accuracy {
            get {
                return this.accuracyField;
            }
            set {
                this.accuracyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<System.DateTime> Timestamp {
            get {
                return this.timestampField;
            }
            set {
                this.timestampField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/location")]
    public partial class CoordinateReferenceSystem {
        
        private string codeField;
        
        private string codeSpaceField;
        
        private string editionField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string code {
            get {
                return this.codeField;
            }
            set {
                this.codeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string codeSpace {
            get {
                return this.codeSpaceField;
            }
            set {
                this.codeSpaceField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string edition {
            get {
                return this.editionField;
            }
            set {
                this.editionField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/common/v2_0")]
    public partial class TimeMetric {
        
        private System.Nullable<TimeMetricsValues> metricField;
        
        private int unitsField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<TimeMetricsValues> Metric {
            get {
                return this.metricField;
            }
            set {
                this.metricField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int Units {
            get {
                return this.unitsField;
            }
            set {
                this.unitsField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/common/v2_0")]
    public enum TimeMetricsValues {
        
        /// <remarks/>
        Millisecond,
        
        /// <remarks/>
        Second,
        
        /// <remarks/>
        Minute,
        
        /// <remarks/>
        Hour,
        
        /// <remarks/>
        Day,
        
        /// <remarks/>
        Week,
        
        /// <remarks/>
        Month,
        
        /// <remarks/>
        Year,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/mms")]
    public partial class MmsMessage {
        
        private string bodyTextField;
        
        private string contentField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string bodyText {
            get {
                return this.bodyTextField;
            }
            set {
                this.bodyTextField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string content {
            get {
                return this.contentField;
            }
            set {
                this.contentField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/mms")]
    public partial class MessageReference {
        
        private string messageIdentifierField;
        
        private string messageServiceActivationNumberField;
        
        private string senderAddressField;
        
        private string subjectField;
        
        private MessagePriority priorityField;
        
        private byte[] messageField;
        
        private System.DateTime dateTimeField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string messageIdentifier {
            get {
                return this.messageIdentifierField;
            }
            set {
                this.messageIdentifierField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string messageServiceActivationNumber {
            get {
                return this.messageServiceActivationNumberField;
            }
            set {
                this.messageServiceActivationNumberField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI")]
        public string senderAddress {
            get {
                return this.senderAddressField;
            }
            set {
                this.senderAddressField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string subject {
            get {
                return this.subjectField;
            }
            set {
                this.subjectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public MessagePriority priority {
            get {
                return this.priorityField;
            }
            set {
                this.priorityField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="base64Binary")]
        public byte[] message {
            get {
                return this.messageField;
            }
            set {
                this.messageField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public System.DateTime dateTime {
            get {
                return this.dateTimeField;
            }
            set {
                this.dateTimeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/mms")]
    public enum MessagePriority {
        
        /// <remarks/>
        Default,
        
        /// <remarks/>
        Low,
        
        /// <remarks/>
        Normal,
        
        /// <remarks/>
        High,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(TypeName="DeliveryInformation", Namespace="http://www.csapi.org/schema/mms")]
    public partial class DeliveryInformation1 {
        
        private string addressField;
        
        private DeliveryStatus1 deliveryStatusField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI")]
        public string address {
            get {
                return this.addressField;
            }
            set {
                this.addressField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public DeliveryStatus1 deliveryStatus {
            get {
                return this.deliveryStatusField;
            }
            set {
                this.deliveryStatusField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(TypeName="DeliveryStatus", Namespace="http://www.csapi.org/schema/mms")]
    public enum DeliveryStatus1 {
        
        /// <remarks/>
        DeliveredToNetwork,
        
        /// <remarks/>
        DeliveryUncertain,
        
        /// <remarks/>
        DeliveryImpossible,
        
        /// <remarks/>
        MessageWaiting,
        
        /// <remarks/>
        DeliveredToTerminal,
        
        /// <remarks/>
        DeliveryNotificationNotSupported,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/sms")]
    public partial class SMSMessage {
        
        private string messageField;
        
        private string smsServiceActivationNumberField;
        
        private string senderAddressField;
        
        private System.Nullable<MessageFormat> messageFormatField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string Message {
            get {
                return this.messageField;
            }
            set {
                this.messageField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string SmsServiceActivationNumber {
            get {
                return this.smsServiceActivationNumberField;
            }
            set {
                this.smsServiceActivationNumberField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string SenderAddress {
            get {
                return this.senderAddressField;
            }
            set {
                this.senderAddressField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<MessageFormat> MessageFormat {
            get {
                return this.messageFormatField;
            }
            set {
                this.messageFormatField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/sms")]
    public enum MessageFormat {
        
        /// <remarks/>
        ASCII,
        
        /// <remarks/>
        UCS2,
        
        /// <remarks/>
        GB18030,
        
        /// <remarks/>
        GB2312,
        
        /// <remarks/>
        Binary,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/sms")]
    public partial class DeliveryInformation {
        
        private string addressField;
        
        private System.Nullable<DeliveryStatus> deliveryStatusField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string Address {
            get {
                return this.addressField;
            }
            set {
                this.addressField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<DeliveryStatus> DeliveryStatus {
            get {
                return this.deliveryStatusField;
            }
            set {
                this.deliveryStatusField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/sms")]
    public enum DeliveryStatus {
        
        /// <remarks/>
        Delivered,
        
        /// <remarks/>
        DeliveryUncertain,
        
        /// <remarks/>
        DeliveryImpossible,
        
        /// <remarks/>
        MessageWaiting,
        
        /// <remarks/>
        DeliveryToTerminal,
        
        /// <remarks/>
        DeliveryNotificationNotSupported,
        
        /// <remarks/>
        KeyWordFilterFailed,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class APRegistrationRsp {
        
        private System.Nullable<APRegResult> regResultField;
        
        private int nextIntervalField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<APRegResult> RegResult {
            get {
                return this.regResultField;
            }
            set {
                this.regResultField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int NextInterval {
            get {
                return this.nextIntervalField;
            }
            set {
                this.nextIntervalField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/ap")]
    public enum APRegResult {
        
        /// <remarks/>
        success,
        
        /// <remarks/>
        illegalAP,
        
        /// <remarks/>
        CMAbilityNotSup,
        
        /// <remarks/>
        repeatedReg,
        
        /// <remarks/>
        svcAddrMismatch,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class APStatusRepReq {
        
        private string aPidField;
        
        private APStatusType aPStatusField;
        
        private int aPPidField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string APid {
            get {
                return this.aPidField;
            }
            set {
                this.aPidField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public APStatusType APStatus {
            get {
                return this.aPStatusField;
            }
            set {
                this.aPStatusField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int APPid {
            get {
                return this.aPPidField;
            }
            set {
                this.aPPidField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/ap")]
    public enum APStatusType {
        
        /// <remarks/>
        Normal,
        
        /// <remarks/>
        OutofActiveTime,
        
        /// <remarks/>
        NeedRegistration,
        
        /// <remarks/>
        OutofService,
        
        /// <remarks/>
        Paused,
        
        /// <remarks/>
        Closed,
        
        /// <remarks/>
        WaitingforConfirm,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class APStatusRepRsp {
        
        private string nextCommandField;
        
        private int nextIntervalField;
        
        private string[] apSvcAuthTypeField;
        
        private string[] apSvcPerfCmdTypeField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string NextCommand {
            get {
                return this.nextCommandField;
            }
            set {
                this.nextCommandField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int NextInterval {
            get {
                return this.nextIntervalField;
            }
            set {
                this.nextIntervalField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ApSvcAuthType", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string[] ApSvcAuthType {
            get {
                return this.apSvcAuthTypeField;
            }
            set {
                this.apSvcAuthTypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ApSvcPerfCmdType", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string[] ApSvcPerfCmdType {
            get {
                return this.apSvcPerfCmdTypeField;
            }
            set {
                this.apSvcPerfCmdTypeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class APLogOutReq {
        
        private string aPidField;
        
        private int aPPidField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string APid {
            get {
                return this.aPidField;
            }
            set {
                this.aPidField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int APPid {
            get {
                return this.aPPidField;
            }
            set {
                this.aPPidField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class APLogOutRsp {
        
        private System.Nullable<APLogoutResult> logoutResultField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<APLogoutResult> LogoutResult {
            get {
                return this.logoutResultField;
            }
            set {
                this.logoutResultField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/ap")]
    public enum APLogoutResult {
        
        /// <remarks/>
        success,
        
        /// <remarks/>
        illegalAPid,
        
        /// <remarks/>
        repeatedLogout,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class PauseAPReq {
        
        private string aPidField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string APid {
            get {
                return this.aPidField;
            }
            set {
                this.aPidField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class PauseAPRsp {
        
        private bool aCKField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public bool ACK {
            get {
                return this.aCKField;
            }
            set {
                this.aCKField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class RecoveryAPReq {
        
        private string aPidField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string APid {
            get {
                return this.aPidField;
            }
            set {
                this.aPidField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class RecoveryAPRsp {
        
        private bool aCKField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public bool ACK {
            get {
                return this.aCKField;
            }
            set {
                this.aCKField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class APSvcAuthenticReq {
        
        private string aPidField;
        
        private string apSvcAuthTypeField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string APid {
            get {
                return this.aPidField;
            }
            set {
                this.aPidField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApSvcAuthType {
            get {
                return this.apSvcAuthTypeField;
            }
            set {
                this.apSvcAuthTypeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class APSvcAuthenticRsp {
        
        private string aPidField;
        
        private string apSvcAuthResultField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string APid {
            get {
                return this.aPidField;
            }
            set {
                this.aPidField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApSvcAuthResult {
            get {
                return this.apSvcAuthResultField;
            }
            set {
                this.apSvcAuthResultField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class APSvcPerfCmdReq {
        
        private string aPidField;
        
        private string apSvcPerfCmdTypeField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string APid {
            get {
                return this.aPidField;
            }
            set {
                this.aPidField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApSvcPerfCmdType {
            get {
                return this.apSvcPerfCmdTypeField;
            }
            set {
                this.apSvcPerfCmdTypeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class AlarmReq {
        
        private string alarmIdField;
        
        private int orgseverityField;
        
        private int orgtypeField;
        
        private string probablecauseField;
        
        private string eventTimeField;
        
        private string ackTimeField;
        
        private string clearTimeField;
        
        private int activestatusField;
        
        private string alarmtitleField;
        
        private string alarmTextField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string alarmId {
            get {
                return this.alarmIdField;
            }
            set {
                this.alarmIdField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int orgseverity {
            get {
                return this.orgseverityField;
            }
            set {
                this.orgseverityField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int orgtype {
            get {
                return this.orgtypeField;
            }
            set {
                this.orgtypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string probablecause {
            get {
                return this.probablecauseField;
            }
            set {
                this.probablecauseField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string eventTime {
            get {
                return this.eventTimeField;
            }
            set {
                this.eventTimeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ackTime {
            get {
                return this.ackTimeField;
            }
            set {
                this.ackTimeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string clearTime {
            get {
                return this.clearTimeField;
            }
            set {
                this.clearTimeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int activestatus {
            get {
                return this.activestatusField;
            }
            set {
                this.activestatusField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string alarmtitle {
            get {
                return this.alarmtitleField;
            }
            set {
                this.alarmtitleField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string alarmText {
            get {
                return this.alarmTextField;
            }
            set {
                this.alarmTextField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class AlarmRsp {
        
        private string recodeField;
        
        private string reMsgField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string recode {
            get {
                return this.recodeField;
            }
            set {
                this.recodeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string reMsg {
            get {
                return this.reMsgField;
            }
            set {
                this.reMsgField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class APSvcPerfReportReq {
        
        private string aPidField;
        
        private string apSvcPerfDataField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string APid {
            get {
                return this.aPidField;
            }
            set {
                this.aPidField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApSvcPerfData {
            get {
                return this.apSvcPerfDataField;
            }
            set {
                this.apSvcPerfDataField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class startNotificationRequest {
        
        private string applicationIdField;
        
        private MessageNotificationType[] messageNotificationField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationId {
            get {
                return this.applicationIdField;
            }
            set {
                this.applicationIdField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("MessageNotification", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public MessageNotificationType[] MessageNotification {
            get {
                return this.messageNotificationField;
            }
            set {
                this.messageNotificationField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class stopNotificationRequest {
        
        private string applicationIdField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationId {
            get {
                return this.applicationIdField;
            }
            set {
                this.applicationIdField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class sendSmsRequest {
        
        private string applicationIDField;
        
        private string[] destinationAddressesField;
        
        private string extendCodeField;
        
        private string messageField;
        
        private System.Nullable<MessageFormat> messageFormatField;
        
        private System.Nullable<SendMethodType> sendMethodField;
        
        private bool deliveryResultRequestField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("DestinationAddresses", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string[] DestinationAddresses {
            get {
                return this.destinationAddressesField;
            }
            set {
                this.destinationAddressesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ExtendCode {
            get {
                return this.extendCodeField;
            }
            set {
                this.extendCodeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string Message {
            get {
                return this.messageField;
            }
            set {
                this.messageField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<MessageFormat> MessageFormat {
            get {
                return this.messageFormatField;
            }
            set {
                this.messageFormatField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<SendMethodType> SendMethod {
            get {
                return this.sendMethodField;
            }
            set {
                this.sendMethodField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public bool DeliveryResultRequest {
            get {
                return this.deliveryResultRequestField;
            }
            set {
                this.deliveryResultRequestField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/sms")]
    public enum SendMethodType {
        
        /// <remarks/>
        Normal,
        
        /// <remarks/>
        Instant,
        
        /// <remarks/>
        Long,
        
        /// <remarks/>
        Structured,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class sendSmsResponse {
        
        private string requestIdentifierField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string RequestIdentifier {
            get {
                return this.requestIdentifierField;
            }
            set {
                this.requestIdentifierField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class notifySmsDeliveryStatusRequest {
        
        private string requestIdentifierField;
        
        private DeliveryInformation[] deliveryInformationField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string RequestIdentifier {
            get {
                return this.requestIdentifierField;
            }
            set {
                this.requestIdentifierField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("DeliveryInformation", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public DeliveryInformation[] DeliveryInformation {
            get {
                return this.deliveryInformationField;
            }
            set {
                this.deliveryInformationField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class GetReceivedSmsRequest {
        
        private string applicationIDField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class GetSmsDeliveryStatusRequest {
        
        private string applicationIDField;
        
        private string requestIdentifierField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string RequestIdentifier {
            get {
                return this.requestIdentifierField;
            }
            set {
                this.requestIdentifierField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class notifySmsReceptionRequest {
        
        private SMSMessage messageField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public SMSMessage Message {
            get {
                return this.messageField;
            }
            set {
                this.messageField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class sendMessageRequest {
        
        private string applicationIDField;
        
        private string[] addressesField;
        
        private string extendCodeField;
        
        private string subjectField;
        
        private MessagePriority priorityField;
        
        private bool receiptRequestField;
        
        private byte[] contentField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("addresses", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI")]
        public string[] addresses {
            get {
                return this.addressesField;
            }
            set {
                this.addressesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ExtendCode {
            get {
                return this.extendCodeField;
            }
            set {
                this.extendCodeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string subject {
            get {
                return this.subjectField;
            }
            set {
                this.subjectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public MessagePriority priority {
            get {
                return this.priorityField;
            }
            set {
                this.priorityField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public bool receiptRequest {
            get {
                return this.receiptRequestField;
            }
            set {
                this.receiptRequestField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="base64Binary", IsNullable=true)]
        public byte[] Content {
            get {
                return this.contentField;
            }
            set {
                this.contentField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class sendMessageResponse {
        
        private string requestIdentifierField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string requestIdentifier {
            get {
                return this.requestIdentifierField;
            }
            set {
                this.requestIdentifierField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class getMessageDeliveryStatusRequest {
        
        private string applicationIDField;
        
        private string requestIdentifierField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string requestIdentifier {
            get {
                return this.requestIdentifierField;
            }
            set {
                this.requestIdentifierField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class getReceivedMessagesRequest {
        
        private string applicationIDField;
        
        private MessagePriority priorityField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public MessagePriority priority {
            get {
                return this.priorityField;
            }
            set {
                this.priorityField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class getMessageRequest {
        
        private string applicationIDField;
        
        private string messageRefIdentifierField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string messageRefIdentifier {
            get {
                return this.messageRefIdentifierField;
            }
            set {
                this.messageRefIdentifierField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class getMessageResponse {
        
        private MmsMessage mmsMessageField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public MmsMessage mmsMessage {
            get {
                return this.mmsMessageField;
            }
            set {
                this.mmsMessageField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class notifyMessageReceptionRequest {
        
        private string applicationIDField;
        
        private MessageReference messageField;
        
        private string contentField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public MessageReference message {
            get {
                return this.messageField;
            }
            set {
                this.messageField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string Content {
            get {
                return this.contentField;
            }
            set {
                this.contentField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class notifyMessageDeliveryReceiptRequest {
        
        private string correlatorField;
        
        private DeliveryInformation1[] deliveryStatusField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string correlator {
            get {
                return this.correlatorField;
            }
            set {
                this.correlatorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("deliveryStatus", Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public DeliveryInformation1[] deliveryStatus {
            get {
                return this.deliveryStatusField;
            }
            set {
                this.deliveryStatusField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class getLocationRequest {
        
        private string requesterField;
        
        private string addressField;
        
        private int requestedAccuracyField;
        
        private int acceptableAccuracyField;
        
        private TimeMetric maximumAgeField;
        
        private TimeMetric responseTimeField;
        
        private System.Nullable<DelayTolerance> toleranceField;
        
        private string applicationIdField;
        
        private ServiceType serviceTypeField;
        
        private CoordinateReferenceSystem crsField;
        
        private LocType locTypeField;
        
        private Priority prioField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string Requester {
            get {
                return this.requesterField;
            }
            set {
                this.requesterField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string Address {
            get {
                return this.addressField;
            }
            set {
                this.addressField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int RequestedAccuracy {
            get {
                return this.requestedAccuracyField;
            }
            set {
                this.requestedAccuracyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int AcceptableAccuracy {
            get {
                return this.acceptableAccuracyField;
            }
            set {
                this.acceptableAccuracyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public TimeMetric MaximumAge {
            get {
                return this.maximumAgeField;
            }
            set {
                this.maximumAgeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public TimeMetric ResponseTime {
            get {
                return this.responseTimeField;
            }
            set {
                this.responseTimeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<DelayTolerance> Tolerance {
            get {
                return this.toleranceField;
            }
            set {
                this.toleranceField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationId {
            get {
                return this.applicationIdField;
            }
            set {
                this.applicationIdField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public ServiceType serviceType {
            get {
                return this.serviceTypeField;
            }
            set {
                this.serviceTypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public CoordinateReferenceSystem crs {
            get {
                return this.crsField;
            }
            set {
                this.crsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public LocType locType {
            get {
                return this.locTypeField;
            }
            set {
                this.locTypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public Priority prio {
            get {
                return this.prioField;
            }
            set {
                this.prioField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/location")]
    public enum DelayTolerance {
        
        /// <remarks/>
        NoDelay,
        
        /// <remarks/>
        LowDelay,
        
        /// <remarks/>
        DelayTolerant,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/location")]
    public enum ServiceType {
        
        /// <remarks/>
        STANDARD,
        
        /// <remarks/>
        EMERGENCY,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/location")]
    public enum LocType {
        
        /// <remarks/>
        CURRENT,
        
        /// <remarks/>
        LAST,
        
        /// <remarks/>
        CURRENT_OR_LAST,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/location")]
    public enum Priority {
        
        /// <remarks/>
        NORMAL,
        
        /// <remarks/>
        HIGH,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class getLocationResponse {
        
        private LocationInfo resultField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public LocationInfo Result {
            get {
                return this.resultField;
            }
            set {
                this.resultField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class getLocationForGroupRequest {
        
        private string requesterField;
        
        private string[] addressesField;
        
        private int requestedAccuracyField;
        
        private int acceptableAccuracyField;
        
        private TimeMetric maximumAgeField;
        
        private TimeMetric responseTimeField;
        
        private System.Nullable<DelayTolerance> toleranceField;
        
        private string applicationIdField;
        
        private ServiceType serviceTypeField;
        
        private CoordinateReferenceSystem crsField;
        
        private LocType locTypeField;
        
        private Priority prioField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string Requester {
            get {
                return this.requesterField;
            }
            set {
                this.requesterField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Addresses", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string[] Addresses {
            get {
                return this.addressesField;
            }
            set {
                this.addressesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int RequestedAccuracy {
            get {
                return this.requestedAccuracyField;
            }
            set {
                this.requestedAccuracyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int AcceptableAccuracy {
            get {
                return this.acceptableAccuracyField;
            }
            set {
                this.acceptableAccuracyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public TimeMetric MaximumAge {
            get {
                return this.maximumAgeField;
            }
            set {
                this.maximumAgeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public TimeMetric ResponseTime {
            get {
                return this.responseTimeField;
            }
            set {
                this.responseTimeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public System.Nullable<DelayTolerance> Tolerance {
            get {
                return this.toleranceField;
            }
            set {
                this.toleranceField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationId {
            get {
                return this.applicationIdField;
            }
            set {
                this.applicationIdField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public ServiceType serviceType {
            get {
                return this.serviceTypeField;
            }
            set {
                this.serviceTypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public CoordinateReferenceSystem crs {
            get {
                return this.crsField;
            }
            set {
                this.crsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public LocType locType {
            get {
                return this.locTypeField;
            }
            set {
                this.locTypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public Priority prio {
            get {
                return this.prioField;
            }
            set {
                this.prioField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class startPeriodicNotificationRequest {
        
        private string requesterField;
        
        private string[] addressesField;
        
        private int requestedAccuracyField;
        
        private TimeMetric frequencyField;
        
        private TimeMetric durationField;
        
        private string applicationIdField;
        
        private ServiceType serviceTypeField;
        
        private CoordinateReferenceSystem crsField;
        
        private LocType locTypeField;
        
        private Priority prioField;
        
        private string eventNotificationField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string Requester {
            get {
                return this.requesterField;
            }
            set {
                this.requesterField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Addresses", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string[] Addresses {
            get {
                return this.addressesField;
            }
            set {
                this.addressesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public int RequestedAccuracy {
            get {
                return this.requestedAccuracyField;
            }
            set {
                this.requestedAccuracyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public TimeMetric Frequency {
            get {
                return this.frequencyField;
            }
            set {
                this.frequencyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public TimeMetric Duration {
            get {
                return this.durationField;
            }
            set {
                this.durationField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationId {
            get {
                return this.applicationIdField;
            }
            set {
                this.applicationIdField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public ServiceType serviceType {
            get {
                return this.serviceTypeField;
            }
            set {
                this.serviceTypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public CoordinateReferenceSystem Crs {
            get {
                return this.crsField;
            }
            set {
                this.crsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public LocType locType {
            get {
                return this.locTypeField;
            }
            set {
                this.locTypeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public Priority prio {
            get {
                return this.prioField;
            }
            set {
                this.prioField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string eventNotification {
            get {
                return this.eventNotificationField;
            }
            set {
                this.eventNotificationField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class startPeriodicNotificationResponse {
        
        private string correlatorField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string correlator {
            get {
                return this.correlatorField;
            }
            set {
                this.correlatorField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class EndNotificationRequest {
        
        private string applicationIdField;
        
        private string correlatorField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationId {
            get {
                return this.applicationIdField;
            }
            set {
                this.applicationIdField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string correlator {
            get {
                return this.correlatorField;
            }
            set {
                this.correlatorField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class LocationNotificationRequest {
        
        private string correlatorField;
        
        private LocationData[] dataField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string correlator {
            get {
                return this.correlatorField;
            }
            set {
                this.correlatorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("data", Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public LocationData[] data {
            get {
                return this.dataField;
            }
            set {
                this.dataField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class LocationErrorRequest {
        
        private string correlatorField;
        
        private string addressField;
        
        private ServiceError reasonField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string correlator {
            get {
                return this.correlatorField;
            }
            set {
                this.correlatorField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI", IsNullable=true)]
        public string Address {
            get {
                return this.addressField;
            }
            set {
                this.addressField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public ServiceError Reason {
            get {
                return this.reasonField;
            }
            set {
                this.reasonField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class LocationEndRequest {
        
        private string correlatorField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string correlator {
            get {
                return this.correlatorField;
            }
            set {
                this.correlatorField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class sendPushRequest {
        
        private string applicationIDField;
        
        private string[] addressesField;
        
        private string targetURLField;
        
        private string extendCodeField;
        
        private string subjectField;
        
        private bool receiptRequestField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("addresses", Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI")]
        public string[] addresses {
            get {
                return this.addressesField;
            }
            set {
                this.addressesField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI")]
        public string targetURL {
            get {
                return this.targetURLField;
            }
            set {
                this.targetURLField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ExtendCode {
            get {
                return this.extendCodeField;
            }
            set {
                this.extendCodeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string subject {
            get {
                return this.subjectField;
            }
            set {
                this.subjectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public bool receiptRequest {
            get {
                return this.receiptRequestField;
            }
            set {
                this.receiptRequestField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class sendPushResponse {
        
        private string requestIdentifierField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string requestIdentifier {
            get {
                return this.requestIdentifierField;
            }
            set {
                this.requestIdentifierField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class getPushDeliveryStatusRequest {
        
        private string applicationIDField;
        
        private string requestIdentifierField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string requestIdentifier {
            get {
                return this.requestIdentifierField;
            }
            set {
                this.requestIdentifierField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class notifyPushDeliveryReceiptRequest {
        
        private string requestIdentifierField;
        
        private DeliveryInformation2[] deliveryStatusField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string requestIdentifier {
            get {
                return this.requestIdentifierField;
            }
            set {
                this.requestIdentifierField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("deliveryStatus", Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public DeliveryInformation2[] deliveryStatus {
            get {
                return this.deliveryStatusField;
            }
            set {
                this.deliveryStatusField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class makeUssdRequest {
        
        private string applicationIDField;
        
        private string destinationAddressField;
        
        private UssdArray ussdMessageField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI")]
        public string destinationAddress {
            get {
                return this.destinationAddressField;
            }
            set {
                this.destinationAddressField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public UssdArray ussdMessage {
            get {
                return this.ussdMessageField;
            }
            set {
                this.ussdMessageField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class makeUssdResponse {
        
        private string ussdIdentifierField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ussdIdentifier {
            get {
                return this.ussdIdentifierField;
            }
            set {
                this.ussdIdentifierField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class handleUssdRequest {
        
        private string ussdIdentifierField;
        
        private string senderAddressField;
        
        private string ussdMessageField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ussdIdentifier {
            get {
                return this.ussdIdentifierField;
            }
            set {
                this.ussdIdentifierField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="anyURI")]
        public string senderAddress {
            get {
                return this.senderAddressField;
            }
            set {
                this.senderAddressField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ussdMessage {
            get {
                return this.ussdMessageField;
            }
            set {
                this.ussdMessageField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class handleUssdResponse {
        
        private UssdArray resultField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public UssdArray result {
            get {
                return this.resultField;
            }
            set {
                this.resultField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class ussdContinueRequest {
        
        private string ussdIdentifierField;
        
        private UssdArray ussdMessageField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ussdIdentifier {
            get {
                return this.ussdIdentifierField;
            }
            set {
                this.ussdIdentifierField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public UssdArray ussdMessage {
            get {
                return this.ussdMessageField;
            }
            set {
                this.ussdMessageField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class ussdContinueResponse {
        
        private UssdArray returnMessageField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public UssdArray returnMessage {
            get {
                return this.returnMessageField;
            }
            set {
                this.returnMessageField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class endUssdRequest {
        
        private string applicationIDField;
        
        private string ussdMessageField;
        
        private string ussdIdentifierField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ApplicationID {
            get {
                return this.applicationIDField;
            }
            set {
                this.applicationIDField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string UssdMessage {
            get {
                return this.ussdMessageField;
            }
            set {
                this.ussdMessageField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string UssdIdentifier {
            get {
                return this.ussdIdentifierField;
            }
            set {
                this.ussdIdentifierField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
    public partial class notifyUssdEndRequest {
        
        private string ussdIdentifierField;
        
        private EndReason endReasonField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ussdIdentifier {
            get {
                return this.ussdIdentifierField;
            }
            set {
                this.ussdIdentifierField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public EndReason endReason {
            get {
                return this.endReasonField;
            }
            set {
                this.endReasonField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://www.csapi.org/schema/ussd")]
    public enum EndReason {
        
        /// <remarks/>
        UserEnd,
        
        /// <remarks/>
        Busy,
        
        /// <remarks/>
        UserAbsent,
        
        /// <remarks/>
        IllegalEquipment,
        
        /// <remarks/>
        SystemError,
        
        /// <remarks/>
        TimeOut,
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void APRegistrationCompletedEventHandler(object sender, APRegistrationCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class APRegistrationCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal APRegistrationCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public APRegistrationRsp Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((APRegistrationRsp)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void APStatusRepCompletedEventHandler(object sender, APStatusRepCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class APStatusRepCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal APStatusRepCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public APStatusRepRsp Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((APStatusRepRsp)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void APLogOutCompletedEventHandler(object sender, APLogOutCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class APLogOutCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal APLogOutCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public APLogOutRsp Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((APLogOutRsp)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void PauseAPCompletedEventHandler(object sender, PauseAPCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class PauseAPCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal PauseAPCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public PauseAPRsp Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((PauseAPRsp)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void RecoveryAPCompletedEventHandler(object sender, RecoveryAPCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class RecoveryAPCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal RecoveryAPCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public RecoveryAPRsp Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((RecoveryAPRsp)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void APSvcAuthenticCompletedEventHandler(object sender, APSvcAuthenticCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class APSvcAuthenticCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal APSvcAuthenticCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public APSvcAuthenticRsp Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((APSvcAuthenticRsp)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void APSvcPerfCmdCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void APSvcAlarmCompletedEventHandler(object sender, APSvcAlarmCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class APSvcAlarmCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal APSvcAlarmCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public AlarmRsp Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((AlarmRsp)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void APSvcPerfReportCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void startNotificationCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void stopNotificationCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void sendSmsCompletedEventHandler(object sender, sendSmsCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class sendSmsCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal sendSmsCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public sendSmsResponse Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((sendSmsResponse)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void notifySmsDeliveryStatusCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void GetReceivedSmsCompletedEventHandler(object sender, GetReceivedSmsCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetReceivedSmsCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetReceivedSmsCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public SMSMessage[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((SMSMessage[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void GetSmsDeliveryStatusCompletedEventHandler(object sender, GetSmsDeliveryStatusCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetSmsDeliveryStatusCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetSmsDeliveryStatusCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public DeliveryInformation[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((DeliveryInformation[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void notifySmsReceptionCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void sendMessageCompletedEventHandler(object sender, sendMessageCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class sendMessageCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal sendMessageCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public sendMessageResponse Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((sendMessageResponse)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void getMessageDeliveryStatusCompletedEventHandler(object sender, getMessageDeliveryStatusCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class getMessageDeliveryStatusCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal getMessageDeliveryStatusCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public DeliveryInformation1[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((DeliveryInformation1[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void getReceivedMessagesCompletedEventHandler(object sender, getReceivedMessagesCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class getReceivedMessagesCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal getReceivedMessagesCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public MessageReference[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((MessageReference[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void getMessageCompletedEventHandler(object sender, getMessageCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class getMessageCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal getMessageCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public getMessageResponse Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((getMessageResponse)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void notifyMessageReceptionCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void notifyMessageDeliveryReceiptCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void getLocationCompletedEventHandler(object sender, getLocationCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class getLocationCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal getLocationCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public getLocationResponse Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((getLocationResponse)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void getLocationForGroupCompletedEventHandler(object sender, getLocationForGroupCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class getLocationForGroupCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal getLocationForGroupCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public LocationData[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((LocationData[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void startPeriodicNotificationCompletedEventHandler(object sender, startPeriodicNotificationCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class startPeriodicNotificationCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal startPeriodicNotificationCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public startPeriodicNotificationResponse Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((startPeriodicNotificationResponse)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void EndNotificationCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void LocationNotificationCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void LocationErrorCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void LocationEndCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void sendPushCompletedEventHandler(object sender, sendPushCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class sendPushCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal sendPushCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public sendPushResponse Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((sendPushResponse)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void getPushDeliveryStatusCompletedEventHandler(object sender, getPushDeliveryStatusCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class getPushDeliveryStatusCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal getPushDeliveryStatusCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public DeliveryInformation2[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((DeliveryInformation2[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void notifyPushDeliveryReceiptCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void makeUssdCompletedEventHandler(object sender, makeUssdCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class makeUssdCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal makeUssdCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public makeUssdResponse Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((makeUssdResponse)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void handleUssdCompletedEventHandler(object sender, handleUssdCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class handleUssdCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal handleUssdCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public handleUssdResponse Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((handleUssdResponse)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void ussdContinueCompletedEventHandler(object sender, ussdContinueCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class ussdContinueCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal ussdContinueCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public ussdContinueResponse Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((ussdContinueResponse)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void endUssdCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")]
    public delegate void notifyUssdEndCompletedEventHandler(object sender, System.ComponentModel.AsyncCompletedEventArgs e);
