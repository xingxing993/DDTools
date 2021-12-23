function newval=ConvertViaType(val,datatype)

try
    if ~ischar(val)
        if length(val)==1 % if not table convert
            newval = eval([datatype '(' num2str(val) ');']);
        else % if table, not convert 
            newval=val;
        end
    else % defined in base workspace
        if isnan(str2double(val))
            newval=evalin('base',[datatype '(' val ');']);
            disp(['###ConvertViaType: ''' val ''' was defined in base workspace!!!']);
        else
            newval=str2double(val);
        end
    end
catch exception
%     newval=0;
%     disp(exception.message);
    newval=0;
    warning(exception.message);
end

end