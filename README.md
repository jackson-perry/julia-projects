# Private-julia-projects
This is a short julia program for simulationg M/M/n queueing models
you need to provide:
average time between arrivals 
average service time
number of minutes to run each simulation
total number of simulations to run
a maiximum number of customers available per simulation
an integer to seed the random number
the program will simulate a customer showing up at a queue randomly according to an exponential distribution with an average tiem as provided.
customers will move to the front of the line are recieve a service for a random exponential itme with a mean provided. 
the simulation will find the mean time in the queue, mean time in service, mean time in system and number of customers served.
it will run this simulation a number of times and create an array for each of: mean time in the queue, mean time in service, 
mean time in system and number of customers served with an entry for each. It will print to the terminal the arveage over all runs of the 
simulation along with 95% confidence interval of each statistic.

It is written in Julia v1.5.3 with the following libraies required:
SimJulia
Distributions
HypothesisTests
Printf
ResumableFunctions
Random
