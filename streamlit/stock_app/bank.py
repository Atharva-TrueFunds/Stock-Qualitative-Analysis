class Account:
    def __init__(self, bal, acc, name):
        self.balance = bal
        self.account_no = acc
        self.name = name
        print(self.name)
        print(self.account_no)

    def debit(self, amount):
        self.balance -= amount
        if self.balance < 0:
            print(
                "Do you want to procced with the balance is going into negative if yes press y and if no press n "
            )
            a = input("Enter your answer: ")
            if a == "y":
                print(self.balance)
            elif a == "n":
                self.balance += amount
                print(self.balance)

    def credit(self, amount):
        self.balance += amount
        if self.balance < 0:
            print(
                "Do you want to procced with the balance is going into negative if yes press y and if no press n "
            )
            a = input("Enter your answer: ")
            if a == "y":
                print(self.balance)
            elif a == "n":
                self.balance -= amount
                print(self.balance)

    def get_balance(self):
        print(self.balance)


acc1 = Account(10000, 12345, "1234")
acc1.debit(1000)
acc1.credit(5000)

acc2 = Account(500, 346565, "Abhi")
acc2.debit(200)
acc2.debit(500)
