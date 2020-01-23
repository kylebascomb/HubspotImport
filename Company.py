

class Company:

    def __init__(self, name, domain):
        self.name = name
        self.domain = domain

    def __eq__(self, other):
        return self.domain == other.domain

    def __hash__(self):
        return hash(self.__repr__())

    def __repr__(self):
        return self.domain

    def __str__(self):
        return self.name, self.domain


