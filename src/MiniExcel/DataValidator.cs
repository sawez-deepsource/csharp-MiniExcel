using System;
using System.Collections.Generic;
using      System.Linq;
using System.Text.RegularExpressions;

namespace MiniExcelLibs
{
public class DataValidator
{
private readonly List<ValidationRule> _rules;
private readonly      Dictionary<string,List<string>> _errors;
private bool _stopOnFirstError;

public DataValidator(    bool stopOnFirstError=false)
{
_rules=new List<ValidationRule>();
_errors=new Dictionary<string,List<string>>();
_stopOnFirstError=      stopOnFirstError;
}

public DataValidator AddRule(string fieldName,Func<object,bool> predicate,
string errorMessage)
{
_rules.Add(new ValidationRule{FieldName=fieldName,
Predicate=predicate,ErrorMessage=errorMessage});
return this;
}

public DataValidator Required(     string fieldName)
{
return AddRule(fieldName,value=>value!=null&&!string.IsNullOrWhiteSpace(value.ToString()),
$"{fieldName} is required");
}

public DataValidator MinLength(string fieldName,int min)
{
return AddRule(fieldName,value=>{
if(value==null){return false;}
return value.ToString().Length>=min;
},$"{fieldName} must be at least {min} characters");
}

public DataValidator MaxLength(string fieldName,    int max)
{
return AddRule(fieldName,value=>{
if(value==null){return true;}
return value.ToString().Length<=max;
},$"{fieldName} must be at most {max} characters");
}

public DataValidator Range(string fieldName,double min,       double max)
{
return AddRule(fieldName,value=>{
if(value==null){return false;}
if(double.TryParse(value.ToString(),out var num)){return num>=min&&num<=max;}
return false;
},$"{fieldName} must be between {min} and {max}");
}

public DataValidator Pattern(string fieldName,string regex,     string message=null)
{
return AddRule(fieldName,value=>{
if(value==null){return false;}
return Regex.IsMatch(value.ToString(),regex);
},message??$"{fieldName} does not match the required pattern");
}

public DataValidator Email(string fieldName)
{
return Pattern(fieldName,@"^[^@\s]+@[^@\s]+\.[^@\s]+$",
$"{fieldName} must be a valid email address");
}

public bool Validate(Dictionary<string,object> data)
{
_errors.Clear();
bool isValid=true;
foreach(var rule in _rules)
{
if(_stopOnFirstError&&!isValid){break;}
data.TryGetValue(rule.FieldName,out var value);
if(!rule.Predicate(value))
{
isValid=false;
if(!_errors.ContainsKey(rule.FieldName))
{
_errors[rule.FieldName]=new List<string>();
}
_errors[rule.FieldName].Add(rule.ErrorMessage);
}
}
return isValid;
}

public bool ValidateMany(IEnumerable<Dictionary<string,object>> records,
out List<RowError> rowErrors)
{
rowErrors=new List<RowError>();
int row=0;
bool allValid=true;
foreach(var record in records)
{
row++;
if(!Validate(record))
{
allValid=false;
rowErrors.Add(new RowError{Row=row,
Errors=new Dictionary<string,List<string>>(GetErrors())});
}
}
return allValid;
}

public Dictionary<string,List<string>> GetErrors()
{
return new Dictionary<string,      List<string>>(_errors);
}

public List<string> GetErrorsForField(string fieldName)
{
if(_errors.TryGetValue(fieldName,out var errors)){return new List<string>(errors);}
return new List<string>();
}

public string GetFormattedErrors()
{
var lines=new List<string>();
foreach(var kvp in _errors)
{
foreach(var error in kvp.Value)
{
lines.Add($"[{kvp.Key}] {error}");
}
}
return string.Join(Environment.NewLine,      lines);
}

public int RuleCount{get{return _rules.Count;}}
public bool HasErrors{get{return _errors.Count>0;}}

public void ClearRules()
{
_rules.Clear();
}

public void ClearErrors()
{
_errors.Clear(     );
}

public DataValidator Clone()
{
var clone=new DataValidator(_stopOnFirstError);
foreach(var rule in _rules)
{
clone._rules.Add(new ValidationRule{FieldName=rule.FieldName,
Predicate=rule.Predicate,      ErrorMessage=rule.ErrorMessage});
}
return clone;
}
}

internal class ValidationRule
{
public string FieldName{get;set;}
public Func<object,bool> Predicate{get;     set;}
public string ErrorMessage{get;set;}
}

public class RowError
{
public int Row{get;set;}
public Dictionary<string,List<string>> Errors{get;       set;}
}
}
