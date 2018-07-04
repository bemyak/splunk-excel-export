Splunk.Module.ExcelExport = $.klass(Splunk.Module, {

    initialize: function($super, container) {
        $super(container);
        this.logger = Splunk.Logger.getLogger("ExcelExport.js");
        this.messenger = Splunk.Messenger.System.getInstance();
        this.popupDiv = $('.excelPopup', this.container).get(0);
        if (Splunk.util.getCurrentView() == 'flashtimeline_excel') {
            this.fixCSS($('#splunk_version', this.container).val());
        }   
        $('#excel_export', this.container).click(this.exportHandler.bind(this));
    },  

    fixCSS: function(version) {
        if (version.indexOf('4.3') == 0) {
            $(this.container).css('background-color', '#EDEDE7');
        } else if (version.indexOf('4.1') == 0 || version.indexOf('4.0') == 0) {
            this.messender.send('warn', 'splunk.search', 'The detected version of Splunk is incompatible with the Excel Export module.  Please upgrade to the latest version of Splunk.'); 
        }   
    },

    exportHandler: function(event) {
        var context = this.getContext();
        var search = context.get('search');
        var sid = search.job.getSearchId();
        if (search.job.isDone()) {
            this.popup = new Splunk.Popup(this.popupDiv, {
                cloneFlag: false,
                title: _("Excel Export"),
                pclass: 'configPopup',
                buttons: [
                    {
                        label: _("Cancel"),
                        type: "secondary",
                        callback: function() {
                            return true; 
                        }.bind(this)
                    },
                    {
                        label: _("Export"),
                        type: "primary",
                        callback: function() {
                            var countstr =  $(exportPopupHandle).find(".exMaxcount").val();
                            var count =  parseInt(countstr, 10);
                            if (isNaN(count) || count<1 || countstr!=count) {
                                alert(_("Must export at least one result"));
                                return false;
                            }   
                            this.messenger.send('info', 'splunk.search', 
                                                _('Your excel file will begin downloading momentarily.  Please do not navigate away from the page.'));
                            return $(exportPopupHandle).find(".exForm").submit();
                        }.bind(this)
                    }
                ]
            });
            exportPopupHandle = this.popup.getPopup();
            var exportForm = $(exportPopupHandle).find('.exForm')[0];
            exportForm.action = Splunk.util.make_url('custom/excel_export/excel/' + sid);
        } else {
            this.messenger.send('info', 'splunk.search', 
                                _('Please wait for your search to complete before exporting to Excel'));
            return false;
        }
    }
});
