import turtle

SCREEN_WIDTH = 600
SCREEN_HEIGHT = 600
TARGET_LLEFT_X = 100
TARGET_LLEFT_Y = 250
TARGET_WIDTH = 25
FORCE_FACTOR= 30
PROJECTILE_SPEED = 1
NORTH = 90
SOUTH = 270
EAST = 0
WEST = 180

turtle.setup(SCREEN_WIDTH, SCREEN_HEIGHT)

turtle.hideturtle()
turtle.speed(0)
turtle.penup()
turtle.goto(TARGET_LLEFT_X, TARGET_LLEFT_Y)
turtle.pendown()
turtle.setheading(EAST)
turtle.forward(TARGET_WIDTH)
turtle.setheading(NORTH)
turtle.forward(TARGET_WIDTH)
turtle.setheading(WEST)
turtle.forward(TARGET_WIDTH)
turtle.setheading(SOUTH)
turtle.forward(TARGET_WIDTH)
turtle.penup()

turtle.goto(0,0)
turtle.setheading(EAST)
turtle.showturtle()
turtle.speed(PROJECTILE_SPEED)

print('Welcome to my game! The objective is to enter an ' + 
        'angle and speed in order for the arrow to hit the target. '
       'You can use decimals. *Hint* You may need to use decimals lol... Good luck!')
angle = float(input("Enter the projectile's angle: "))
force = float(input("Enter the launch force (1-10): "))

distance = force * FORCE_FACTOR

turtle.setheading (angle)

turtle.pendown()
turtle.forward(distance)

if (turtle.xcor() >= TARGET_LLEFT_X and
    (TARGET_LLEFT_X + TARGET_WIDTH) and
    turtle.ycor() >= TARGET_LLEFT_Y and
    turtle.ycor() <= (TARGET_LLEFT_Y + TARGET_WIDTH)):
        print("Direct Hit!")
else:
    print('Target missed, try again.')

turtle.done()