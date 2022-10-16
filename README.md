# Production-Monitoring

The script for following production parts (stators and rotors) on current processes steps. The script doesn't count scrap or NOK parts and finished parts as well.
It checks the control card (google sheet) and fulfillment of field (Part Acceptance for next process). If it is NOK it doesn't go further and stops - the parts is not counted. If it is OK it goes further to first not fulfilled field - it means that the part is in this process now and script puts it to file with part reference number.
A great help for Production Management Team.
