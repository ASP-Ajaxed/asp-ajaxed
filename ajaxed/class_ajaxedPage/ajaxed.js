function ajaxed() {}

ajaxed.prototype.indicator = Element.extend(document.createElement('div')).addClassName('ajaxLoadingIndicator');

//registers the loading indicator
Ajax.Responders.register({
	onCreate: function() {
		var s = ajaxed.prototype.indicator.style;
		s.position = 'absolute';
		s.top = ajaxed.getWindowScroll().top + 'px';
		if (Ajax.activeRequestCount == 1) document.body.appendChild(ajaxed.prototype.indicator);
	},
	onComplete: function() {
		if (Ajax.activeRequestCount == 0) document.body.removeChild(ajaxed.prototype.indicator);
	}
});

//optional: onComplete, url (because of bug in iis5 http://support.microsoft.com/kb/216493)
ajaxed.callback = function(theAction, func, params, onComplete, url) {
	if (params) {
		params = $H(params);
	} else {
		if ($('frm')) {
			params = $H($('frm').serialize(true));
		} else {
			params = new Hash();
		}
	}
	params = params.merge({PageAjaxed: theAction});
	uri = window.location.href;
	if (ajaxed.prototype.debug) ajaxed.debug("Action (to be handled in callback):\n\n" + theAction);
	if (uri.endsWith('/') && url) uri += url;
	if (ajaxed.prototype.debug) ajaxed.debug("Params passed to xhr:\n\n" + params.toJSON());
	new Ajax.Request(uri, {
		method: 'post',
		parameters: params.toQueryString(),
		requestHeaders: {Accept: 'application/json'},
		onSuccess: function(trans) {
			if (ajaxed.prototype.debug) ajaxed.debug("Response on callback:\n\n" + trans.responseText);
			if (!trans.responseText.startsWith('{ "root":')) {
				ajaxed.callbackFailure(trans);
			} else {
				if (func) func(trans.responseText.evalJSON(true).root);
			}
		},
		onFailure: ajaxed.callbackFailure,
		onComplete: onComplete
	}); 
}
ajaxed.callbackFailure = function(transport) {
	friendlyMsg = transport.responseText;
	friendlyMsg = friendlyMsg.replace(new RegExp("(<head[\\s\\S]*?</head>)|(<script[\\s\\S]*?</script>)", "gi"), "");
	friendlyMsg = friendlyMsg.stripTags();
	friendlyMsg = friendlyMsg.replace(new RegExp("[\\s]+", "gi"), " ");
	alert(friendlyMsg);
}
ajaxed.debug = function(msg){
	alert("<DEBUG MESSAGE>\n\n" + msg);
}

// Returns the location and the size of the square
// which is shown to the user (windowscroll)
// Core code from - quirksmode.org
ajaxed.getWindowScroll = function(parent){
	var T, L, W, H;
	parent = parent || document.body;
	if (parent != document.body) {
		T = parent.scrollTop;
		L = parent.scrollLeft;
		W = parent.scrollWidth;
		H = parent.scrollHeight;
	} else {
		var w = window;
		with (w.document) {
			if (w.document.documentElement && documentElement.scrollTop) {
				T = documentElement.scrollTop;
				L = documentElement.scrollLeft;
			} else {
				if (w.document.body) {
					T = body.scrollTop;
					L = body.scrollLeft;
				}
			}
			if (w.innerWidth) {
				W = w.innerWidth;
				H = w.innerHeight;
			} else {
				if (w.document.documentElement && documentElement.clientWidth) {
					W = documentElement.clientWidth;
					H = documentElement.clientHeight;
				} else {
					W = body.offsetWidth;
					H = body.offsetHeight
				}
			}
		}
	}
	return {
		top: T,
		left: L,
		width: W,
		height: H
	};
}