##Goal
Write a script that reads this input file. contact_inaccurate_fields/input_files/subscriptions_contact_map_fields.xlsx

For every matching Stripe Customer Email in sheet 1 = Export For Stripe Subs Field Update, find the corresponding Email in sheet = Export contact field update.

From the sheet Export For Stripe Subs Field Update

Map the values from columns Billing Start Date	Billing End Date	Status	Products	Discount	Coupon 

into these columns in sheet 2 Export For Contact Field Update
Billing Start Date	Billing End Date	Status	Products	Discount	Coupon 







##Context

Here's the conditional 

if NO mulitple Id found for match email map specified columns over.

ELSE 

In sheet one, there are emails that contain multiple Stripe IDs

ex: 
Stripe Email = sratner@umich.edu
Row 21 contains Id = sub_1RarfwKNMTeBGk8yC1pb0qFD
Row 9639 contains Id = sub_1SZzLfKNMTeBGk8yXycfGgvu
Row 9691 contains Id = sub_1SZzKZKNMTeBGk8yitLlIzpb

Must always use the LATEST date in Most Recent Create Date

Therefore Row 9639 contains Id = sub_1SZzLfKNMTeBGk8yXycfGgvu has  Most Recent Create Date of 12/2/2025 14:34:00 

This is the correct matched email to copy the values from.




##Inputs
Xlsx = contact_inaccurate_fields/input_files/subscriptions_contact_map_fields.xlsx

Sheet 1 = Export For Stripe Subs Field Up

Sheet 2 = Export For Contact Field Update

Columns in Sheet 1 = Record ID	Id	Stripe Customer Email	Most Recent Create Date	Billing Start Date	Billing End Date	Status	Products	Discount	Coupon

Columns in Sheet 2 = Email	Billing Start Date	Billing End Date	Status	Products	Discount	Coupon

##Output format

An ouputted file in the new xlsx file named contacts_with_updated_fields


