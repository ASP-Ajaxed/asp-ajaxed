var watcher = null;
function scrollToFloor(el) {
	$(el).scrollTop = $(el).scrollHeight;
}
function loadContent(file, sender) {
	if (watcher) watcher.stop();
	new Ajax.Updater('htmlcontent', file, {evalScripts:true});
	last = $($('htmlcontent').readAttribute('title'));
	if (last) last.removeClassName('active');
	$(sender).addClassName('active');
	$('htmlcontent').writeAttribute('title', sender.id);
}
