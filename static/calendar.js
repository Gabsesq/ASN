;window.calendar = (function(){

	/**
	 * Helper function to create an element with the the given attributes
	 * @param {string} tagname
	 * @param {object} name-value pairs to be set as attributes
	 * @returns {HTMLElement}
	 */
	function makeEle(name, attrs){
		var ele = document.createElement(name);
		for(var i in attrs)
			if(attrs.hasOwnProperty(i))
				ele.setAttribute(i, attrs[i]);
		return ele;
	}
	
	/**
	 * Remove an event handler from an element
	 * @param {HTMLElement} el - element with an event listener to remove
	 * @param {string} type - the type of event being handled (eg, click)
	 * @param {function} handler - the event handler to be removed
	 * @returns {undefined}
	 */
	function removeEventListener(el, type, handler) {
	    if (el.detachEvent) el.detachEvent('on'+type, handler); 
		else el.removeEventListener(type, handler);
	}
	
	/**
	 * Attach an event handler to an element
	 * @param {HTMLElement} el - element to attach an event listener to
	 * @param {string} type - the type of event to listen for (eg, click)
	 * @param {function} handler - the event handler to be attached
	 * @returns {undefined}
	 */
	function addEventListener(el, type, handler) {
	    if (el.attachEvent) el.attachEvent('on'+type, handler); 
		else el.addEventListener(type, handler);
	}
	
	class Calendar{
		
		constructor(elem, opts){
			opts = opts || {};
			
			this._eventGroups = [];
			this.selectedDates = [];
			this.elem = elem;
			this.disabledDates = (opts.disabledDates || []).map(d=>this.formatDateMMDDYY(d));
			this.abbrDay = opts.hasOwnProperty('abbrDay') ? opts.abbrDay : true;
			this.abbrMonth = opts.hasOwnProperty('abbrMonth') ? opts.abbrMonth : true;
			this.abbrYear = opts.hasOwnProperty('abbrYear') ? opts.abbrYear : true;
			this.onDayClick = opts.onDayClick || function(){};
			this.onEventClick = opts.onEventClick || function(){};
			this.onMonthChanged = opts.onMonthChanged || function(){};
			this.beforeDraw = (opts.beforeDraw || function(){}).bind(this);
			this.events = [];
			this.month = opts.hasOwnProperty('month') ? opts.month-1 : (new Date()).getMonth();
			this.year = opts.hasOwnProperty('year') ? opts.year : (new Date()).getFullYear();
			this.ellipse = opts.hasOwnProperty('ellipse') ? opts.ellipse : true;
			this.daysOfWeekFull = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
			this.daysOfWeekAbbr = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
			this.monthsFull = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
			this.monthsAbbr = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
			
			opts.events = opts.events || [];
			for(var i=opts.events.length; i--;) this.addEvent(opts.events[i], false);
			
			this.elem.classList.add("CalendarJS");
			
			// Built-in event handlers
			this._loadNextMonth = this.loadNextMonth.bind(this);
			this._loadPrevMonth = this.loadPreviousMonth.bind(this);
			
			this._dayClicked = function(e){
				var dataEvents = e.target.getAttribute('data-events');
				if (!dataEvents) return; // Skip if no data-events attribute
				
				var evtids = dataEvents.split(",");
				var evts = [];
				for(var i=0; i<evtids.length; i++) 
					if(this.events[evtids[i]]!==undefined) 
						evts.push(this.events[evtids[i]]);
				var date = new Date(this.year, this.month, e.target.getAttribute('data-day'));
				this.onDayClick.call(this, date, evts);
			}.bind(this);
			
			this._eventClicked = function(e){
				var target = e.target.tagName === 'SPAN' ? e.target.parentElement : e.target;
				var evtid = target.getAttribute('data-eventid');
				this.onEventClick.call(this, this.events[evtid]);
			}.bind(this);
			
			this.drawCalendar();
		}
		
		/**
		 * Add a date to the disabled dates list
		 * @param {type} date
		 * @returns {undefined}
		 */
		disableDate(date){
			var formatted = this.formatDateMMDDYY(date);
			if(!this.disabledDates.includes(formatted)){
				this.disabledDates.push(formatted);
			}
		}
		
		/**
		 * Given a Date object, format and return date string in format MM/DD/YYYY (without leading zeroes)
		 * @param {Date} date
		 * @returns {String}
		 */
		formatDateMMDDYY(date){
			return (date.getMonth() + 1) + "/" +
					date.getDate() + "/" +
					date.getFullYear();
		}
		
		/**
		 * Add an event to the calendar
		 * @param {object} evt
		 * @returns {Calendar instance}
		 */
		async addEvent(evt, draw=true){
			var hasDesc = (evt.hasOwnProperty("desc")),
				hasDate = (evt.hasOwnProperty("date")),
				hasStartDate = (evt.hasOwnProperty("startDate")),
				hasEndDate = (evt.hasOwnProperty("endDate"));
			if(hasDesc && (hasDate  || (hasStartDate && hasEndDate))){
				if(hasStartDate && hasEndDate){
					if(+evt.startDate <= +evt.endDate) this.events.push(evt);
					else throw new Error("Start date must occur before end date.");
				}else this.events.push(evt);
			}else throw new Error("All events must have a 'desc' property and either a 'date' property or a 'startDate' and 'endDate' property");
			if(draw) await this.drawCalendar();
			return this;
		}
		
		/**
		 * Load the next month to the calendar
		 * @returns {Calendar instance}
		 */
		async loadNextMonth () {
			this.month = this.month + 1 > 11 ? 0 : this.month + 1;
			if (this.month === 0) this.year++;
			await this.drawCalendar();
			this.onMonthChanged.call(this, this.month+1, this.year);
			return this;
		}
		
		/**
		 * Load the previous month to the calendar
		 * @returns {Calendar instance}
		 */
		async loadPreviousMonth(){
			this.month = this.month - 1 > -1 ? this.month - 1 : 11;
			if(this.month===11) this.year--;
			await this.drawCalendar();
			this.onMonthChanged.call(this, this.month+1, this.year);
			return this;
		}
		
		/**
		 * Get a list of eventts in a certain time range
		 * @param {date} date1
		 * @param {date} date2
		 * @returns {Array of events in range}
		 */
		getEventsDuring(date1, date2){
			if(undefined === date2) date2 = date1;
			var lowdate = +date1>+date2 ? date2 : date1;
			var highdate = +date1>+date2 ? date1 : date2;

			var morn = +(new Date(lowdate.getFullYear(), lowdate.getMonth(), lowdate.getDate(), 0, 0, 0));
			var night = +(new Date(highdate.getFullYear(), highdate.getMonth(), highdate.getDate(), 23, 59, 59));

			var result = [];
			for(let i=0; i<this.events.length; i++){
				if(this.events[i].hasOwnProperty("startDate")){
					var eventStart = +this.events[i].startDate;
					var eventEnd = +this.events[i].endDate;
				}else{
					var eventStart = +this.events[i].date;
					var eventEnd = +this.events[i].date;
				}
				var startsToday = (eventStart>=morn && eventStart<=night);
				var endsToday = (eventEnd>=morn && eventEnd<=night);
				var continuesToday = (eventStart<morn && eventEnd>night);

				if(startsToday || endsToday || continuesToday) result.push(this.events[i]);
			}
			return result;
		}
		
		/**
		 * Clear the selected dates
		 * @returns {Calendar instance}
		 */
		clearSelection(){
			var active = this.elem.getElementsByClassName("cjs-active");
			for(let i=active.length; i--;) active[i].classList.remove('cjs-active');
			this.selectedDates = [];
			return this;
		}
		
		/**
		 * Select a date from teh calendar
		 * @param {date} date
		 * @returns {Calendar instance}
		 */
		selectDate(date){
			if(this.month !== date.getMonth()) return;
			if(this.year !== date.getFullYear()) return;
			this.elem.getElementsByClassName("cjs-dayCell"+(date.getDate()))[0]
				.parentNode.parentNode.parentNode.classList.add("cjs-active");
			this.selectedDates.push({
				day: date.getDate(),
				month: date.getMonth(),
				year: date.getFullYear()
			});
			return this;
		}
		
		/**
		 * Select a range of dates
		 * @param {date} date1
		 * @param {date} date2
		 * @returns {Calendar instance}
		 */
		selectDateRange(date1, date2){
			if(this.month !== date1.getMonth()) return;
			if(this.year !== date1.getFullYear()) return;
			if(this.month !== date2.getMonth()) return;
			if(this.year !== date2.getFullYear()) return;
			var lowdate = +date1>+date2 ? date2.getDate() : date1.getDate();
			var highdate = +date1>+date2 ? date1.getDate() : date2.getDate();
			for(let i=lowdate; i<=highdate; i++)
				this.selectDate(new Date(this.year, this.month, i));
			return this;
		}
		
		/**
		 * Get an array of dates that are selected
		 * @returns {Array of dates}
		 */
		getSelection(){
			var sel = [];
			for(let i=0; i<this.selectedDates.length; i++) 
				sel.push(new Date(this.selectedDates[i].year, this.selectedDates[i].month, this.selectedDates[i].day));
			return sel;
		}
		
		/**
		 * Get the current month being displayed
		 * @returns {object}
		 */
		getCurrentMonth(){
			return {
				month: this.abbrMonth ? this.monthsAbbr[this.month] : this.monthsFull[this.month],
				year: this.year
			};
		}
		
		/**
		 * Redraw the calendar. This is the only function that should be used to update the calendar.
		 * @returns {undefined}
		 */
		async drawCalendar(){
			this.elem.innerHTML = "";
			this.beforeDraw();
			
			// Header
			var header = makeEle("DIV", {"class":"cjs-calHeader"});
			var pMonth = makeEle("SPAN", {"class":"cjs-lastLink cjs-left cjs-bottom-left"});
			var nMonth = makeEle("SPAN", {"class":"cjs-nextLink cjs-right cjs-bottom-right"});
			var cMonth = makeEle("SPAN", {"class":"cjs-moTitle"});
			pMonth.innerHTML = "&lt;";
			nMonth.innerHTML = "&gt;";
			cMonth.innerHTML = (this.abbrMonth ? this.monthsAbbr[this.month] : this.monthsFull[this.month]) + " " + this.year;
			
			// remove old event listeners
			removeEventListener(pMonth, "click", this._loadPrevMonth);
			removeEventListener(nMonth, "click", this._loadNextMonth);
			
			addEventListener(pMonth, "click", this._loadPrevMonth);
			addEventListener(nMonth, "click", this._loadNextMonth);
			
			header.appendChild(pMonth);
			header.appendChild(cMonth);
			header.appendChild(nMonth);
			this.elem.appendChild(header);
			this.elem.appendChild(makeEle("DIV", {"class": "cjs-clearfix"}));
			
			
			// Day headers
			var dows = this.abbrDay ? this.daysOfWeekAbbr : this.daysOfWeekFull;
			var dhead = makeEle("DIV", {"class":"cjs-dayHeader"});
			for(var i=0; i<7; i++){
				var d = makeEle("DIV", {"class": "cjs-dayHeaderCell"});
				d.innerHTML = dows[i];
				dhead.appendChild(d);
			}
			this.elem.appendChild(dhead);
			this.elem.appendChild(makeEle("DIV", {"class": "cjs-clearfix"}));
			
			
			// Calendar days
			var fDay = (new Date(this.year, this.month, 1)).getDay();
			var mLen = (new Date(this.year, this.month+1, 0)).getDate();
			var pLen = (new Date(this.year, this.month, 0)).getDate();
			
			var calBody = makeEle("DIV");
			var row = makeEle("DIV", {"class": "cjs-weekRow"});
			for(var i=1; i<=fDay; i++){
				var day = makeEle("DIV", {"class":"cjs-dayCol cjs-blankday"});
				row.appendChild(day);
			}
			
			var dayCounter = fDay + 1;
			for(var i=1; i<=mLen; i++){
				var isToday = (this.month===(new Date()).getMonth() && this.year===(new Date()).getFullYear() && i===(new Date()).getDate());
				var cls = "cjs-dayCol cjs-calDay" + (isToday?" cjs-today":"");
				var day = makeEle("DIV", {"class":cls, "data-day":i});
				
				var content = makeEle("DIV", {"class":"cjs-dayContent"});
				var table = makeEle("TABLE", {"class":"cjs-dayTable"});
				var tbody = makeEle("TBODY");
				var tr = makeEle("TR");
				
				var td = makeEle("TD", {"class":"cjs-dayCell cjs-dayCell"+i, "data-day": i});
				var dayNum = makeEle("DIV", {"class":"cjs-dateLabel"});
				dayNum.innerHTML = i;
				td.appendChild(dayNum);
				
				// Handle Events
				var morn = +(new Date(this.year, this.month, i, 0,0,0));
				var night = +(new Date(this.year, this.month, i, 23,59,59));
				var events = [];
				var eventids = [];
				for(var n=0; n<this.events.length; n++){
					var eventStart = +this.events[n].date;
					var eventEnd = this.events[n].hasOwnProperty("endDate") ? +this.events[n].endDate : eventStart;
					if((eventStart>=morn && eventStart<=night) || (eventEnd>=morn && eventEnd<=night) || (eventStart<morn && eventEnd>night)){
						events.push(this.events[n]);
						eventids.push(n);
					}
				}
				
				if(events.length>0){
					day.setAttribute("data-events", eventids.join(","));
					addEventListener(day, "click", this._dayClicked);
				}
				
				var eventList = makeEle("DIV");
				for(var n=0; n<events.length; n++){
					var eventObj = events[n];
					var evt = makeEle("DIV", {
						"data-eventid": eventids[n],
						"class": "cjs-calEvent"
					});
					
					// Add tooltip functionality
					evt.setAttribute("title", eventObj.desc);
					
					if(this.ellipse){
						var desc = makeEle("SPAN");
						desc.innerHTML = eventObj.desc;
						evt.appendChild(desc);
						evt.classList.add("cjs-ellipse");
					}else evt.innerHTML = eventObj.desc;
					
					if (eventObj.hasOwnProperty("cls") && eventObj.cls) {
						evt.classList.add(eventObj.cls);
					}
					
					// --- Add delete button ---
					var delBtn = makeEle("SPAN", {"class": "cjs-delBtn"});
					delBtn.innerHTML = "&times;";
					delBtn.style.marginLeft = "6px";
					delBtn.style.cursor = "pointer";
					delBtn.onclick = (e) => {
						e.stopPropagation();
						// Remove event from localStorage by id (fallback to desc+date if no id)
						let eventId = eventObj.id || (eventObj.desc + '|' + eventObj.date.toISOString());
						let savedEvents = JSON.parse(localStorage.getItem('shippingEvents')) || [];
						let updatedEvents = savedEvents.filter(ev => {
							let evId = ev.id || (ev.desc + '|' + new Date(ev.date).toISOString());
							return evId !== eventId;
						});
						localStorage.setItem('shippingEvents', JSON.stringify(updatedEvents));
						location.reload();
					};
					evt.appendChild(delBtn);
					// --- End delete button ---

					// --- Show details on event click ---
					addEventListener(evt, "click", (e) => {
						if (e.target.classList.contains('cjs-delBtn')) return;
						let details = `Details:\n\nDescription: ${eventObj.desc}`;
						if (eventObj.company) details += `\nCompany: ${eventObj.company}`;
						if (eventObj.po) details += `\nPO: ${eventObj.po}`;
						details += `\nDate: ${eventObj.date instanceof Date ? eventObj.date.toLocaleDateString() : eventObj.date}`;
						alert(details);
					});
					// --- End show details ---

					addEventListener(evt, "click", this._eventClicked);
					eventList.appendChild(evt);
				}
				td.appendChild(eventList);
				tr.appendChild(td);
				tbody.appendChild(tr);
				table.appendChild(tbody);
				content.appendChild(table);
				day.appendChild(content);
				
				row.appendChild(day);

				if(dayCounter%7===0){
					this.elem.appendChild(row);
					this.elem.appendChild(makeEle("DIV", {"class": "cjs-clearfix"}));
					row = makeEle("DIV", {"class": "cjs-weekRow"});
				}
				dayCounter++;
			}
			
			var finalDay = (new Date(this.year, this.month, mLen)).getDay();
			for(var i=1; i<7-finalDay; i++){
				var day = makeEle("DIV", {"class":"cjs-dayCol cjs-blankday"});
				row.appendChild(day);
			}
			
			this.elem.appendChild(row);
			this.elem.appendChild(makeEle("DIV", {"class": "cjs-clearfix"}));
		}
	}
	
	return Calendar;
})(); 