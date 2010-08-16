/* Turkish initialisation for the jQuery UI date picker plugin. */
/* Written by Izzet Emre Erkan (kara@karalamalar.net). */
jQuery(function($){
	$.datepicker.regional['tr'] = {
		closeText: 'kapat',
		prevText: '&#x3c;geri',
		nextText: 'ileri&#x3e',
		currentText: 'bugün',
		monthNames: ['Ocak','Þubat','Mart','Nisan','Mayýs','Haziran',
		'Temmuz','Aðustos','Eylül','Ekim','Kasým','Aralýk'],
		monthNamesShort: ['Oca','Þub','Mar','Nis','May','Haz',
		'Tem','Aðu','Eyl','Eki','Kas','Ara'],
		dayNames: ['Pazar','Pazartesi','Salý','Çarþamba','Perþembe','Cuma','Cumartesi'],
		dayNamesShort: ['Pz','Pt','Sa','Ça','Pe','Cu','Ct'],
		dayNamesMin: ['Pz','Pt','Sa','Ça','Pe','Cu','Ct'],
		dateFormat: 'dd.mm.yy', firstDay: 1,
		isRTL: false};
	$.datepicker.setDefaults($.datepicker.regional['tr']);
});