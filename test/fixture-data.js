/* Fake tracker used by make-fixture.js and run-tests.js. All names are invented.
 * JS Date values become real Excel date cells; strings stay text cells. */
const D = (y, m, d) => new Date(y, m - 1, d);
module.exports = {
  fileName: 'L2_L3_Tracker - September 2026.xlsx',
  sheets: {
    TaskList: [
      { TASK: 'Fix login error', Type: 'Incident', Application: 'ONEWAY', Date: D(2026, 9, 2), POC: 'Rami Haddad', Shadowing: 'Lina Saab', Status: 'In Progress', Domain: 'shipping ', 'Ticket ID': 'INC0012345' },
      { TASK: 'Deploy release 4.2', Type: 'Change Request', Application: 'oneway', Date: '05/09/2026', POC: 'HADDAD Rami', Shadowing: '', Status: 'Completed', Domain: 'Shipping', 'Ticket ID': 'CHG0003001' },
      { TASK: 'Config update', Type: 'Change Request', Application: 'Hook', Date: D(2026, 9, 10), POC: 'Nadim', Shadowing: '', Status: 'Pending', Domain: 'LineOps', 'Ticket ID': 'CHG0003002' },
      { TASK: 'Website banner bug', Type: 'Bug', Application: 'Corporate Website', Date: '2026-09-01', POC: 'PagerDuty Connector', Shadowing: 'Lina', Status: 'Requested', Domain: 'COMMUNICATION', 'Ticket ID': 'INC0012346 / INC0012347' },
      { TASK: 'Certificate renewal', Type: 'Request', Application: 'Cert Portal', Date: D(2026, 9, 15), POC: 'DigiCert', Shadowing: 'Nadim Khoury', Status: 'Incomplete', Domain: '', 'Ticket ID': '' },
      { TASK: 'Tissue paper order', Type: 'Request', Application: 'MYEUP', Date: '31/09/2026', POC: 'Lina Saab', Shadowing: '', Status: '', Domain: 'Shipping', 'Ticket ID': 'RITM0005001; HIP-12345' },
      { TASK: 'Data fix', Type: 'Issue', Application: 'Terminal Link', Date: D(2026, 9, 20), POC: 'Lina Saab, Rami Haddad', Shadowing: '', Status: 'Done', Domain: 'Communication', 'Ticket ID': 'INC0012345' },
      { TASK: 'Mobile app outage', Type: 'Incident', Application: 'ONEWAY', Date: '2026-08-25', POC: 'Group Application Security', Shadowing: 'Rami', Status: 'Pending', Domain: 'Shipping', 'Ticket ID': 'INC0012400' },
      { TASK: 'Newsletter template', Type: 'Enhancement', Application: '', Date: D(2026, 9, 28), POC: '', Shadowing: '', Status: 'New', Domain: '', 'Ticket ID': 'HIP-2026/77' },
      { TASK: 'Release notes', Type: 'Deployment', Application: 'Terminal Link', Date: D(2026, 9, 29), POC: 'Nadim Khoury', Shadowing: '', Status: 'Deployed', Domain: 'Communication', 'Ticket ID': 'CHG0003003' },
    ],
    AppDomain: [
      { Application: 'Corporate Website', Domain: 'Communication' },
      { Application: 'Terminal Link', Domain: 'Communication' },
      { Application: 'ONEWAY', Domain: 'Shipping' },
      { Application: 'HOOK', Domain: 'LineOps' },
      { Application: 'MYEUP', Domain: 'MyEup' },
    ],
    Team: [
      { Name: 'Rami HADDAD', Aliases: 'Rami; R. Haddad' },
      { Name: 'Lina SAAB', Aliases: 'Lina' },
      { Name: 'Nadim KHOURY', Aliases: 'Nadim' },
    ],
  },
};
