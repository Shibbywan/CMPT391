# CMPT391
Hotel management system that involves room management, booking systems, and more features
http://cs.wellesley.edu/~cs304/lectures/15-Entity-Relationships/
http://ugweb.cs.ualberta.ca/~c291/W07/exams/solmt04b1.pdf
http://ugweb.cs.ualberta.ca/~c291/W07/lab-slides/ER/ER.pdf
http://candordeveloper.com/2013/01/08/creating-a-sql-server-database-project-in-visual-studio-2012/ great reference


*Have user permissions such that only certain people can access some elements in the database. This will require a method to add employee to database.*

Features in Hotel Management System:
Booking:

Create a method that lets an employee book a customer. The screen (1) will show a list of rooms with attributes floor, number of beds, availability, and smoking. If you click a specific room, it should tell you more information (2) such as if the room contains a Murphy bed, number of bathrooms, the appliances the room has, the measurement of square feet, TV size (?), if it contains a bunk bed, and more (will add more as I see fit). On both Screen 1 and 2 there will be a book button that allows an employee to book a room for a customer. This will open a screen (3) that takes in Customer information such as first name, last name, phone number, number of habitants, date of arrival, date of departure, method of payment, customer preferences, promotions or any coupons (add more as I see fit). There should also be some distinction of paying in person, and making a reservation. There should also be mutual exclusion support such that the same room can not be booked multiple times simultaneously.

Inventory:

Create a method that lets an employee add/remove/update inventory. The screen (4) will show a list of inventory with attributes name, amount, and number. There buttons that allow an employee to add or edit products. Add will bring up a new screen (5) that allows an employee to enter information on a product such as: name, number, supplier, ISBN, description, and cost (ill add more as I see fit). Clicking on a table entry in (4) should open up (5). Screen (5) will also be the same for editing a product, however some fields will be uneditable. 

(Some Ideas: when a product is empty, the row in the table should be highlighted red. If a product is, the row in the table should be highlighted green.)

Human Resources:

Create method that lets an employee manage all employees working at the hotel. The screen (6) will show a list of employees with attributes first name, last name, SIN, employee ID, phone number(s), emergency contact, position, and pay (will add more as see fit).

(note, it seems this will just be a table listing results from a query that gets all employees from Employee IsA Person)

Person:

Create TABLE that shows all people at the hotel, employees and customers. This will be used to manage employee discounts, and other things like retrieving information on past customers, or past employees. This will make it easier for returning customers. Person should contain attributes First Name, Last Name, SIN, Phone Number, Address, and Emergency Contact (will add more). Some of these fields will only be required for employees, some for customers.

