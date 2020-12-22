from win32com.client import Dispatch
import random
from playsound import playsound
random_generator = random.randint(1,3)
def speak(str):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str)


rock = '''
    _______
---'   ____)
      (_____)
      (_____)
      (____)
---.__(___)
'''

paper = '''
    _______
---'   ____)____
          ______)
          _______)
         _______)
---.__________)
'''

scissors = '''
    _______
---'   ____)____
          ______)
       __________)
      (____)
---.__(___)
'''


user_input = int(input("What do you want to choose  \n1) rock press 1 \n2) paper press 2 \n3) scissors press 3\n"))
# random_generator = random.randint(1,3)
if user_input == 1 and random_generator == 2:
    print("You Choose")
    print(rock)

    if __name__ == '__main__':
        with open("paper.txt") as f:
            for item in f.readlines():
                speak(item)
    #print("Computer Choose")
    print(paper)
    playsound('D:\Youtube\lossse.mp3')
    if __name__ == '__main__':
        with open("you_lose.txt") as f:
            for item in f.readlines():
                speak(item)
    #print("Computer Win!")
    print(paper)


elif user_input == 2 and random_generator == 1:
    print("You Choose")
    print(paper)
    #print("Computer Choose")
    if __name__ == '__main__':
        with open("rock.txt") as f:
            for item in f.readlines():
                speak(item)
    print(rock)
    playsound('D:\Youtube\Winning-sound-effect.mp3')
    if __name__ == '__main__':
        with open("you_win.txt") as f:
            for item in f.readlines():
                speak(item)

    #you win!
    print(paper)
elif user_input == 2 and random_generator == 3:
    print("You Choose")
    print(paper)

    if __name__ == '__main__':
        with open("scissors.txt") as f:
            for item in f.readlines():
                speak(item)
    #print("Computer Choose")
    print(scissors)
    playsound('D:\Youtube\lossse.mp3')
    if __name__ == '__main__':
        with open("you_lose.txt") as f:
            for item in f.readlines():
                speak(item)
    #print("Computer Win!")
    print(scissors)
elif user_input == 3 and random_generator == 2:
    print("You Choose")
    print(scissors)
    
    if __name__ == '__main__':
        with open("paper.txt") as f:
            for item in f.readlines():
                speak(item)
    #print("Computer Choose")
    print(paper)
    playsound('D:\Youtube\Winning-sound-effect.mp3')
    if __name__ == '__main__':
        with open("you_win.txt") as f:
            for item in f.readlines():
                speak(item)

    print("You Win!")
    print(scissors)    
elif user_input == 1 and random_generator == 3:
    print("You Choose")
    print(rock)
    
    if __name__ == '__main__':
        with open("scissors.txt") as f:
            for item in f.readlines():
                speak(item)
    # print("Computer Choose")
    print(scissors)
    playsound('D:\Youtube\Winning-sound-effect.mp3')
    if __name__ == '__main__':
        with open("you_win.txt") as f:
            for item in f.readlines():
                speak(item)

    print("You Win!")
    print(rock)
elif user_input == 3 and random_generator == 1:
    print("You Choose")
    print(scissors)
    
    if __name__ == '__main__':
        with open("rock.txt") as f:
            for item in f.readlines():
                speak(item)
    #print("Computer Choose")
    print(rock)
    playsound('D:\Youtube\lossse.mp3')
    if __name__ == '__main__':
        with open("you_lose.txt") as f:
            for item in f.readlines():
                speak(item)
    #print("Computer Win!")
    print(rock)        

print('=========================================================')    