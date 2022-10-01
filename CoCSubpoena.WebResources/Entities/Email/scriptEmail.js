var Email = Email || {};
Email = {
    onLoad: function (exeContext) {
        var Subpeona;
        var _formContext = exeContext.getFormContext();
        var _regardingobject = _formContext.getAttribute("regardingobjectid").getValue();
        if (_regardingobject != null) {
            var _regardingObjectId = _regardingobject[0].id.toString().replace('{','').replace('}','');
            if (_regardingobject[0].entityType != "coc_subpoena")
                return;
            var _query = '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">' +
                '<entity name="coc_subpoena">' +
                '<attribute name="coc_name" />' +
                '<attribute name="coc_casename" />' +
                '<attribute name="coc_casenumber" />' +
                '<attribute name="coc_subpoenanumber" />' +
                '<attribute name="coc_subpoenaid" />' +
                '<filter type="and">' +
                '<condition attribute="coc_subpoenaid" operator="eq" value="' + _regardingObjectId + '" />' +
                '</filter>' +
                '</entity>' +
                '</fetch>';
            _query = encodeURIComponent(_query);
            var _globalContext = Xrm.Utility.getGlobalContext();
            var _clientUrl = _globalContext.getClientUrl();

            var _queryPath = "/api/data/v9.1/coc_subpoenas?fetchXml=" + _query;
            var _requestPath = _clientUrl + _queryPath;
            var _req = new XMLHttpRequest();
            _req.open("GET", _requestPath, false);
            _req.setRequestHeader("Accept", "application/json");
            _req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            _req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    _req.onreadystatechange = null;
                    if (this.status === 200) {
                        var _response = JSON.parse(this.response);
                        if (_response.value.length > 0) {
                            Subpeona = _response.value;
                            _formContext.getAttribute("subject").setValue(Subpeona[0].coc_name);
                        }
                    }
                }
            }
            _req.send();
        }
    }
}