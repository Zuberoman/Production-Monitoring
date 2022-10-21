# Production-Parts-Monitoring

The script for following the current process step of production parts (stators and rotors) by Google Apps Script. 
It checks the control card (google sheet) and fulfillment of field (Part Acceptance for next process). If it is NOK it doesn't go further and stops - the parts is not counted. If it is OK it goes further to first not fulfilled field - it means that the part is in this process now and script puts it to Production Monitoringfile with part reference number and process that is on-going.
A great help for Production Management Team.
