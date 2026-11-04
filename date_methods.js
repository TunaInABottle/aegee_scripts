
///////////////////////////
///// DATE PROTOTYPES /////
///////////////////////////

/**
 * Used for testing
 */
function test_date_proto(){
  Logger.log(NOW)
  //console.log(NOW.addDays(7))
  //console.log(NOW.addWorkDays(28))

  test_weekends();
}

function test_weekends() {
  let date_saturday = new Date("2022-12-31")
  let date_sunday = new Date("2023-01-01")
  
  Logger.log( "Weekend tests" )
  Logger.log( "Saturday: " + date_saturday + " " + date_saturday.getDay())
  Logger.log( date_saturday.isWeekend() )
  Logger.log( "skipping_weekend: " + date_saturday.skipWeekend() )
  Logger.log( `add ${WEEKDAYS_FOLLOWUP} working days equals ${date_saturday.addWorkDays(WEEKDAYS_FOLLOWUP)}`)

  Logger.log( "Sunday: " + date_sunday )
  Logger.log( date_sunday.isWeekend() )
  Logger.log( "skipping_weekend: " + date_sunday.skipWeekend() )
  Logger.log( `add ${WEEKDAYS_FOLLOWUP} working days equals ${date_sunday.addWorkDays(WEEKDAYS_FOLLOWUP)}`)

  Logger.log( `added workdays: sunday equal sunday?  ${date_sunday.isAfter(date_saturday)}`)
}


/**
 * Add days to this date
 * 
 * @param {int} days - How many days to add to this
 * @return {Date} date
 */
Date.prototype.addDays = function(days) {
    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}

/**
 * Add days to this Date by considering only workdays
 * 
 * considers Saturday and Sunday as weekends
 * 
 * @param {int} workdays - workdays to add
 * @return {Date} date
 */
Date.prototype.addWorkDays = function(workdays) {
    var old_date = new Date(this.valueOf());
    var new_date = new Date(old_date.valueOf());

    // move away from weekend (STARTING DATE)
    if (new_date.isWeekend()) {
      new_date = new_date.skipWeekend();
      workdays = workdays - 1;
    }

    // add days normally
    new_date = new_date.addDays(workdays)

    // add weekends when workdays is larger than 7
    new_date = new_date.addSkippedWeekends(workdays)

    // skip festivities
    //TODO

    // move away from weekend (AFTER SUMS)
    if (new_date.isWeekend()) {
      new_date = new_date.skipWeekend();
    }
    return new_date;
}

/**
 * check if this happens weekend (Saturday or Sunday)
 * 
 * @return {bool} - yes if this date is Weekend
 */
Date.prototype.isWeekend = function() {
  return (this.getDay() == 0 || this.getDay() == 6);
}

/**
 * Move this Date to a workday
 * 
 * @return {Date} - date schifted to a workday
 */
Date.prototype.skipWeekend = function() {
    let date = new Date(this.valueOf());

    if (date.getDay() == 0) { // SUNDAY, add one day before calc
      date = date.addDays(1);
    } else if (date.getDay() == 6) {// SATURDAY, add two days before calc
      date = date.addDays(2);
    }
    return date;
}

/**
 * add days equal to the number of weekend days to this date if more than a week has been added
 * 
 * Saturday and Sunday are considered weekends
 * 
 * @param {int} days - number of days to add to this
 * @return {Date} - a date where all the weekends have been added
 */
Date.prototype.addSkippedWeekends = function(days) {
    let date = new Date(this.valueOf());
    let n_weekends = Math.floor(days/7);

    while (n_weekends > 0) { // necessary as the number of weekends might do new weekends
      let weekend_days = n_weekends * 2;
      date = date.addDays(weekend_days);
      n_weekends = Math.floor(weekend_days/7);
    }
    date.skipWeekend();
    return date;
}

/**
 * does this date come before "that_date"?
 * 
 * @param {Date} that_date - date to be compared
 * @return {bool} - true if this comes before
 */
Date.prototype.isBefore = function(that) {
  this_date = this.stripTime();
  that_date = that.stripTime();

  return this_date.getTime() < that_date.getTime();
}

/**
 * does this date come after "that_date"?
 * 
 * @param {Date} that_date - date to be compared
 * @return {bool} - true if this comes after
 */
Date.prototype.isAfter = function(that) {
  this_date = this.stripTime();
  that_date = that.stripTime();

  return this_date.getTime() > that_date.getTime();
}

/**
 * does this date come after "that_date"?
 * 
 * @param {Date} that - date to be compared
 * @return {bool} - true if this and "that_date" are the same
 */
Date.prototype.isEqual = function(that) {
  this_date = this.stripTime();
  that_date = that.stripTime();

  return this_date.getTime() === that_date.getTime();
}

/**
 * remove time from this Date
 * 
 * @return {Date} - an object with only year, month, day
 */
Date.prototype.stripTime = function() {
  return new Date(this.getFullYear(), this.getMonth(), this.getDate());
}