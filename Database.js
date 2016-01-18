var Database = (function() {
	var adOpenForwardOnly = 0;
	var adOpenStatic = 3;
	var adLockReadOnly = 1;
	var adLockOptimistic = 3;
	var adVarChar = 200;
	var adParamInput = 1;
    var ObjectCreator = getObjectCreator();
	
    function __db(connectionString) {
        this.ConnectionString = connectionString;
    }
    
    __db.prototype.Query = function(sql) {
        var queryParameters = getQueryParameters(arguments);
        var command = createAdodbCommand(sql, queryParameters, this.ConnectionString);
        return createRecordSet(command, true);
    }
    
    __db.prototype.CreateUpdatableRecordSet = function(sql) {
        var queryParameters = getQueryParameters(arguments);
        var command = createAdodbCommand(sql, queryParameters, this.ConnectionString);
        return createRecordSet(command, false);
    }
    
    __db.prototype.ExecSql = function(sql){
        var queryParameters = getQueryParameters(arguments);
        var command = createAdodbCommand(sql, queryParameters, this.ConnectionString);
        return command.Execute();
    }

	function getQueryParameters(args) {
		var parameters = new Array();
		
		for(var i = 1; i < args.length; i++) {
			parameters[i-1] = args[i];
		}
		
		return parameters;
	}
	
	function createAdodbCommand(sql, parameters, connectionString){
		var queryCommand = ObjectCreator.CreateObject('ADODB.Command');
		queryCommand.ActiveConnection = connectionString;
		queryCommand.CommandText = sql;
		
		setCommandParameters(queryCommand, parameters);
		
		return queryCommand;
	}
	
	function setCommandParameters(queryCommand, parameters) {
		var queryParameter;
		for(var i = 0; i < parameters.length; i++){
			queryParameter = queryCommand.CreateParameter('', adVarChar, adParamInput, 8000, parameters[i]);
			queryCommand.Parameters.Append(queryParameter);
		}
	}

	function createRecordSet(command, readOnly) {
		var recordSet = ObjectCreator.CreateObject('ADODB.Recordset');
		configureCursorAndLockType(recordSet, readOnly);
		recordSet.Open(command);
		return recordSet;
	}
	
	function configureCursorAndLockType(recordSet, readOnly) {
		if(readOnly) {
			recordSet.CursorType = adOpenForwardOnly;
			recordSet.LockType = adLockReadOnly;
		} else {
			recordSet.CursorType = adOpenStatic;
			recordSet.LockType = adLockOptimistic;
		}
	}
    
    function getObjectCreator() {
        if(typeof(Server) === 'object') {
            return Server;
        }else if(typeof(WScript) === 'object') {
            return WScript;
        }else{
            return null;
        }
    }
    
    function createDatabase(connectionString, objectCreator) {
        if(objectCreator) {
            ObjectCreator = objectCreator;
        }
        return new __db(connectionString);
    }
    
    if (typeof(exports) === 'object') {
        exports.Database = createDatabase;
    }
    
    return createDatabase;
})();