// Function to apply the formula to the spreadsheet cells
function ApplyFormula(cell, formulae) {
	var value = 0;
	var formulas = new Array();

	//We can have multiple formulae input so we need to first
	//seperate into individual formulas
	if (formulae.search(/\|\|/) > 0) {
		formulas = formulae.split(/\|\|/);
		//alert("formula0 " + formulas[0]);
	} else {
		formulas[0] = formulae;
	}

	// Now parse the formula to get the different components
	for (var i = 0; i < formulas.length; i++) {
		//alert("formula is " + formulas[i]);
		evaluateFormula(formulas[i]);
	}
}
// Class of range variables. Since the range is a string we want to
// pass the instance and manipulate in the function when evaluated
function RangeVariable() {
	this.value = "";
}

// Evaluate the formula
function evaluateFormula(formula) {
	// Loop through the formula parsing each entity
	var index = 0;
	
	// Split the formula and parse the RHS expression
	var lrformula = formula.split("=");
	var rformula = lrformula[1];

	// Get the ID of the LHS
	var lhstemp = lrformula[0].split(/\[/)[1];
	var lhsId = lhstemp.split(/\]/)[0];

	// Replace the function(...) with function(Array(...)) in the
	// expression
	rformula = replaceFuncWithArr(rformula);
	
	// Loop through the expression and replace the variables with
	// their values. In the case of a range replace with an instance
	// of the RangeVariable class with the value the range string
	var rangeCount = 0;
	var rv = "rangeVariable";
	var rangeVariable = new Array();
	var loopC = 0;
	var retArr = new Array();
	while (1) {
		// Potentially a variable (ignore the rangeVariable 
		// instance array)
		if (rformula[index].match(/\[/)) {
			if (index-rv.length >= 0) {
				if (! rformula.slice(index-rv.length, index).match(rv)) {
					//alert("have a match!");
					retArr = getVariable(rformula, 
							index, 
							rangeCount);
					rformula = retArr[0];
					rangeCount = retArr[2];
					if (retArr[1].length > 0) {
						rangeVar = new RangeVariable();
						rangeVar.value = retArr[1];
						rangeVariable.push(rangeVar);
					}
					// Set the index to 0 as we have resized the
					// expression
					index = 0;
				} else {
					index += 1;
				}
			} else {
				//alert("have another match!");
				retArr = getVariable(rformula, 
						index, 
						rangeCount);
				rformula = retArr[0];
				rangeCount = retArr[2];
				if (retArr[1].length > 0) {
					rangeVar = new RangeVariable();
					rangeVar.value = retArr[1];
					rangeVariable.push(rangeVar);
				}
				// Set the index to 0 as we have resized the
				// expression
				index = 0;
			}	
		} else {
			index += 1;
		}
			if (index == rformula.length) {
			//alert("expression is: " + rformula);
			break;
		}
	}
	var result;
	// Evaluate the expression
		try  {
		result = eval(rformula);
		} catch (err) {
			result =     rformula;
		}
	//alert("lhs id is " + lhsId);
	document.getElementById(lhsId).value = result;

	// Need to cause the onchange event for the target cell to fire
	// in case the target is part of another formula
	if (typeof document.getElementById(lhsId).onchange == 'function') {
		document.getElementById(lhsId).onchange();
	}
	//alert("result is: " + result);
}

// Function to replace the arguments in a function in an expression with
// Array(...). Needed to correctly evaluate the expression
function replaceFuncWithArr(formula) {
	var idx = 0;
	var retArr = new Array();
	rlen = formula.length;
	while (1) {
		if (formula[idx].match(/\(/) &&
				idx-1 > 0 &&
				formula[idx-1].match(/[a-zA-Z0-9]/)) {
			//alert("we have a match " + formula + " " +
			//			formula[idx-1] +
			//		" index " + idx);
			retArr = replaceFunction(formula, idx);
			formula = retArr[0];
			idx = retArr[1];
			rlen = retArr[2];
		} else {
			idx += 1;
		}
		if (idx == rlen) {
			break;
		}
	}
	return formula;
}

// Function to replace the arguments of a function with an array
function replaceFunction(expression, idx) {
	var newExpression = "";
	var bArrayExpr = "\(Array\(";
	var eArrayExpr = "\)\)";
	
	for (var i = idx; i < expression.length; i++) {
		if (expression[i].match(/\)/)) {
			break;
		}
	}
	var partExpr0 = expression.slice(0,idx);
	var partExpr1 = expression.slice(idx,i+1);
	var partExpr2 = expression.slice(i+1);
	var newPart = partExpr1.replace(expression[i], eArrayExpr);
	newPart = newPart.replace(expression[idx], bArrayExpr);
	newExpression = partExpr0 + newPart + partExpr2;
	//alert("newexpr " + newExpression);
	return [newExpression, i + bArrayExpr.length-2 
		+ eArrayExpr.length-2, newExpression.length];
}

// Function to get the variable from the expression
function getVariable(expression, idx, rangeCount) {
	var i = 0;
	var aRange = 0;
	var newExpression = "";
	var newVar = "";
	var range = "";
	var rangeName = "rangeVariable";
	for (i = idx; i < expression.length; i++) {
		// Flag if we have a ':' as that denotes a range
		if (expression[i].match("\:")) {
			aRange = 1;
		}
		if (expression[i].match("\]")) {
			//alert("i " + i + " il " + (i-rangeName.length-1) + " aRange " + aRange);
			if (i-rangeName.length-1 >= 0) {
				if (! expression.slice(i-rangeName.length-1, i+1).match(rangeName)) {
					//alert("have a break!");
					break;
				}
			} else {
				break;
			}
		}
	}
	if (aRange == 1) {
		var newName = rangeName + "[" + rangeCount + "]";
		newExpression = expression.replace(
				expression.slice(idx, i + 1), newName);
		range = expression.slice(idx + 1, i);
		rangeCount += 1;
		//alert("expression is " + expression + " new expression is " + newExpression);
	} else {
		if (document.getElementById(expression.slice(idx+1, i))) { 
			newVar = document.getElementById(
					expression.slice(idx+1, i)).value; 
		} else {
			newVar = 0;
		}
		if (expression[i+1] == ";") {
			expression = expression.replace(expression[i+1], "\,");
		}
		if (!isNumber(newVar))
			newVar='"'  + newVar + '"';
		newExpression = expression.replace(expression.slice(idx, i+1), 
				newVar);
	}
	//alert("expression is: " + newExpression);
	return [newExpression, range, rangeCount];
}

// Function to return an array of data within the range
function processRange(limit) {
	var rangeVals = new Array();
	var limits = limit.split("\:");
	var cnodes = document.getElementsByTagName("input");
	upperN = limits[1].split(/\.[a-zA-Z]+/);
	lowerN = limits[0].split(/\.[a-zA-Z]+/);
	upperCh = limits[1].split(/\d+/);
	lowerCh = limits[0].split(/\d+/);
	var upperR = parseInt(upperN[1]);
	var lowerR = parseInt(lowerN[1]);
	var upperC = upperCh[0];
	var lowerC = lowerCh[0];
	
	//alert("upperR " + upperR + " lowerR " + lowerR);
	//alert("upperC " + upperC + " lowerC " + lowerC);

	for (var i = 0; i < cnodes.length; i++) {
		var id = cnodes[i].getAttribute("id");
		var ids = id.split(/\.[a-zA-Z]+/);
		var idR = parseInt((ids[1]));
		var idC = id.split(/\d+/)[0];
		// alert("name is " + id + " value " + cnodes[i].value);
		// If same row then use normal test
		if ((idR == lowerR && idC >= lowerC && idC <= upperC) ||
				(idR == upperR && idC >= lowerC &&
				 idC <= upperC)) {
			rangeVals.push(parseFloat(cnodes[i].value));
		}
		// If not the same row we have to check within the limits
		// and that the cell is within the bounding box
		if ((idR > lowerR && idR < upperR) &&
				(idC >= lowerC && idC <= upperC)) {
			rangeVals.push(parseFloat(cnodes[i].value));
		}
	}
	//alert("rangeVals is " + rangeVals);
	return rangeVals;
}
function isNumber (o) {
	  return ! isNaN (o-0);
	}

function CONCATENATE(data) {
	//alert("we are called! data.length " + data.length);
	var total = "";
	var vals = new Array();
	for (var i = 0; i < data.length; i++) {
		if (data[i] instanceof Object && data[i].value.match("\:")) {
			vals = processRange(data[i].value);
			for (var j = 0; j < vals.length; j++) {
				total += vals[j];
			}
			//alert("val total is " + total);
		} else {
			//alert("data i is " + data[i]);
			total = total + data[i];
			//alert("total is " + total);
		}
	}
	return total;
}

// Function to calculate the sum of input variables
function SUM(data) {
	//alert("we are called! data.length " + data.length);
	var total = 0;
	var vals = new Array();
	for (var i = 0; i < data.length; i++) {
		if (data[i] instanceof Object && data[i].value.match("\:")) {
			vals = processRange(data[i].value);
			for (var j = 0; j < vals.length; j++) {
				total += vals[j];
			}
			//alert("val total is " + total);
		} else {
			//alert("data i is " + data[i]);
			total = total + data[i];
			//alert("total is " + total);
		}
	}
	return total;
}
