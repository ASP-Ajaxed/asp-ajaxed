function AxdDT(tID, sorted) {
	
	this.table = $(tID);
	this.sorted = '';
	this.lastSelectedRow;
	this.fullsearchQuery = '';
	this.page = 1;
	
	this.goTo = function(pageNr) {
		var t = this;
		this.callback('goTo', {axd_dt_page: pageNr}, function(trans) {
			t.page = pageNr;
		});
	}
	
	this.setSorted = function(col) {
		this.sorted = col.toLowerCase();
		$$('#' + this.table.id + ' th').each(function(el) {
			el.removeClassName("sortedASC").removeClassName("sortedDESC");
		})
		this.sorted.split(',').each(function(c) {
			c = c.strip();
			var name = c.replace(/ desc$| asc$/ig, '').strip();
			var dir = "ASC";
			if (c.match(/ desc$/i)) dir = "DESC";
			thCol = $$("#" + this.table.id + ' th.axdCol_' + name);
			if (thCol.length > 0) thCol[0].addClassName('sorted' + dir);
		}, this);
	}
	this.setSorted(sorted);
	
	this.callback = function(method, params, onCompleted, container) {
		//we pass all other criterias as well, so that everything is remembered
		//on the callback the actual params are merged and thus the existing one overriden
		var allParams = $H({
			axd_dt_id: this.table.id,
			axd_dt_sort: this.sorted,
			axd_dt_fullsearch: this.fullsearchQuery,
			axd_dt_page: this.page
		});
		//also add the record selections
		var selected = [];
		$$('#' + this.table.id + " .axdDTColSelection input:checked").each(function(el){
			selected.push(el.value);
		})
		allParams.set(this.table.id, selected);
		var t = this;
		ajaxed.callback(
			"axd_dt_" + method, 
			(container) ? container : this.table.id + "_body", 
			allParams.merge(params), 
			function(trans) {
				//the trick with paging is that the last row of the body is the paging row
				//we take it and move it into the footer.
				$$('#' + t.table.id + ' tfoot')[0].update($$('#' + t.table.id + '_body > tr:last-child')[0]);
				onCompleted();
			}
		);
	}
	
	this.toggleRow = function(rowID, selected, unselectLast) {
		var row = $(rowID);
		var css = "axdDTRowSelected"
		if (unselectLast && this.lastSelectedRow) this.lastSelectedRow.removeClassName(css);
		if (selected) row.addClassName(css)
		else row.removeClassName(css);
		this.lastSelectedRow = row;
	}
	
	this.sort = function(col) {
		var dir = "ASC";
		col = col.strip().toLowerCase();
		this.sorted.split(',').each(function(c) {
			c = c.strip().toLowerCase();
			var name = c.replace(/ desc$| asc$/ig, '').toLowerCase();
			if (name == col) {
				if (c.match(/ desc$/i)) dir = "ASC"
				else if (c.match(/ asc$/i)) dir = "DESC"
				else dir = "DESC";
			}
		});
		var sort = col + ' ' + dir;
		var t = this;
		this.callback('sort', {axd_dt_sort: sort}, function(trans) {
			t.setSorted(sort);
		});
	}
	
	this.search = function(query) {
		var t = this;
		//when fullsearch is used we need to reset it to page 1.
		this.page = 1;
		this.callback('fullsearch', {axd_dt_fullsearch: query}, function(trans) {
			t.fullsearchQuery = query;
		});
	}
	
}
