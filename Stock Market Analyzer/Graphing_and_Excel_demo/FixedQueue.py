class FixedQueue(object):
    """description of class"""
    def __init__(self, max_length = 10):
        self.array = [None]*max_length
        self.front = self.rear = self.occupancy = 0
  
    def enqueue(self, data):
        if not self.isFull():
            self.array[self.rear] = data
            if self.rear + 1 >= len(self.array):
                self.rear = 0
            else:
                self.rear += 1
                self.occupancy += 1
        else:
            print("\nQueue is full")
  
    def dequeue(self):
        if not self.isEmpty():
            x = self.array[self.front]
            self.array[self.front] = None
            if self.front + 1 >= len(self.array):
                self.front = 0
            else:
                self.front += 1
                self.occupancy -= 1
            return x
        else:
            print("\nQueue is empty")

    def peek(self):
        if not self.isEmpty():
            return self.array[self.front]
        else:
            print("\nQueue is empty")

    def getArray(self):
        return self.array

    def isFull(self):
        return self.occupancy >= len(self.array)

    def isEmpty(self):
        return self.occupancy == 0

