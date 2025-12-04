# Hotfix Invoice Changed

Due to a change in the billing system, which replaced the installation number with the consumer unit number, the billing system began to malfunction.

The installation number was used to verify if the generated bill corresponded to the requested installation. Without this information, the system rejected all generated bills.

To resolve the problem, I had to collect the consumer unit number information and send it to the TEL_BOT program so that it would check the consumer unit instead of the installation number.

Therefore, it was a double change, both in SAP_BOT and TEL_BOT. Because of these changes, it is no longer mandatory to request a bill by installation number; now the order or meter numbers can be used as before.
