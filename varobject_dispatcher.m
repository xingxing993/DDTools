function varobject_dispatcher(varobject, context_info, verbose)

% context_info: current context information, 
% with string fields <file>, <sheet>, <var>, and integer field <row> 

if nargin<3
    verbose = true;
end

% print warning if data object already exists
if verbose && evalin('base',sprintf('exist(''%s'',''var'')',varobject.Name)) %if data object exists in base, prints warning
    fprintf('## Warning: %s has been overwritten in workspace (Sheet: %s, Line: %u)\n',context_info.var,context_info.sheet,context_info.row);
end
% NV
if ~isempty(regexp(context_info.sheet,'^NV')) %sheet name starts with NV
    if verbose&&varobject.Name(1)~='N'
        fprintf('## Warning: %s doesn''t begin with "N", check if necessary (Sheet: %s, Line: %u)\n',context_info.var,context_info.sheet,context_info.row);
    end
    create_NVsignal(varobject);
    
% MEASURE
elseif ~isempty(regexp(context_info.sheet,'^MEASURE')) %sheet name starts with MEASURE
    if verbose&&isempty(regexp(varobject.Name,'^[VSDR]'))
        fprintf('## Warning: %s doesn''t begin with "V/S/D/R", check if necessary (Sheet: %s, Line: %u)\n',context_info.var,context_info.sheet,context_info.row);
    end
    create_signal(varobject);
    
% CALIRBRATION
elseif ~isempty(regexp(context_info.sheet,'^CALIBRATION')) %sheet name starts with CALIBRATION
    if verbose&&isempty(regexp(varobject.Name,'^[KMA]'))
        fprintf('## Warning: %s doesn''t begin with "K/A/M", check if necessary (Sheet: %s, Line: %u)\n',context_info.var,context_info.sheet,context_info.row);
    end
    % project dependent calibratable feature
    if isfield(varobject, context_info.project)
        varobject.Calibratable = varobject.(context_info.project);
    end
    create_parameter(varobject);
else
    
end