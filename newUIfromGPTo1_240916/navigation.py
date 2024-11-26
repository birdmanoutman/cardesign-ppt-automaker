# navigation.py
class Navigation:
    def __init__(self):
        self.history = []
        self.index = -1

    def navigate_to(self, path):
        if self.index < len(self.history) - 1:
            self.history = self.history[:self.index + 1]
        self.history.append(path)
        self.index += 1

    def back(self):
        if self.index > 0:
            self.index -= 1
            return self.history[self.index]
        return None

    def forward(self):
        if self.index < len(self.history) - 1:
            self.index += 1
            return self.history[self.index]
        return None

    def current_path(self):
        if self.index >= 0:
            return self.history[self.index]
        return None

    def can_go_back(self):
        return self.index > 0

    def can_go_forward(self):
        return self.index < len(self.history) - 1
