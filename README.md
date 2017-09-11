BlueCare Configuration Deploy
Requirements: Powershell 2.0, Powershell 3.0 for Powershell Remoting

Script features:
- support infinite BlueCare configuration files without modification of the main script
- duplicated servers will be check and cause script to break before any further actions
- source files will be check if they exist and if they has valid data
- for each column name, each source file will be read only once and place in global memory dataset
- overwrite operation for all servers will be performed at once with the data from global memory dataset
- modern error/exception handling, log format easy to import for other tools
- messages on screen represents actual state of the operations and their results
