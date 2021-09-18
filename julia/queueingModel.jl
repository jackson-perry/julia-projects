using SimJulia
using Distributions
using HypothesisTests
using Printf
using ResumableFunctions
using Random

#set inputs
runs =100
max_customer =150
seed = 123
theta_arrival = 7
theta_service = 5.0
servers = 1
simulation_time = 500
#initialize arrays
arrival_time_Array=Float64[]
departure_time_Array=Float64[]
service_time_Array=Float64[]
time_in_system_Array=Float64[]
time_in_queue_Array=Float64[]
aggregate_time_in_service_Array=Float64[]
aggregate_time_in_queue_Array=Float64[]
aggregate_time_in_system_Array=Float64[]
customers_served=Int[]


#simulation logic
function get_arrival_time()
    inter_arrival = Exponential(theta_arrival)
    push!(arrival_time_Array,rand(inter_arrival))
    return sum(arrival_time_Array)
end

@resumable function queue(env::Environment,name::Int, server::Resource,theta_service::Float64)
    @yield timeout(env, get_arrival_time())
    arrive = now(env)
    #println(name, " arrive at ", arrive)
    @yield request(server)
    served = now(env)
    #println(name, " served at ",served)
    service_time = rand(Exponential(theta_service))
    push!(service_time_Array,service_time)
    @yield timeout(env, service_time)
    depart=now(env)
    push!(departure_time_Array,depart)
    #println(name, " leaving system at ",depart)
    time_in_system=depart-arrive
    push!(time_in_system_Array,time_in_system)
    time_in_queue= time_in_system-service_time
    push!(time_in_queue_Array,time_in_queue)
    @yield release(server)
end
#
for i in 1:runs
    Random.seed!(seed)
    arrival_time_Array=Float64[]
    departure_time_Array=Float64[]
    service_time_Array=Float64[]
    time_in_system_Array=Float64[]
    time_in_queue_Array=Float64[]
    queue
    sim= Simulation()
    server = Resource(sim,servers)
    for i in 1:max_customer
        @process queue(sim, i, server, theta_service)
    end
    run(sim,simulation_time)
    seed=seed+1
#create aggregates
  
    mean_time_in_service=mean(service_time_Array)
    mean_time_in_system=mean(time_in_system_Array)
    mean_time_in_queue=mean(time_in_queue_Array)
    push!(customers_served,length(departure_time_Array))
    #println("Mean time in service ",mean_time_in_service)
    #println("Mean time in system ",mean_time_in_system)
    #println("Mean time in queue ", mean_time_in_queue)
    #println("Customers served ", customers_served)
    push!(aggregate_time_in_queue_Array,mean_time_in_queue)
    push!(aggregate_time_in_service_Array,mean_time_in_service)
    push!(aggregate_time_in_system_Array,mean_time_in_system)
end
println("Runs: ", runs, "  Sim time: ", simulation_time, "  Seed: ", seed)
println("CAUTION max customers is: ", max_customer, "  ensure it is well over customers served\n")

println("Mean time in service ", mean(aggregate_time_in_service_Array))
println("95% confidecn interval",ci(OneSampleTTest(aggregate_time_in_service_Array)),"\n")

println("Mean time in system ", mean(aggregate_time_in_system_Array))
println("95% confidecn interval",ci(OneSampleTTest(aggregate_time_in_system_Array)),"\n")

println("Mean time in queue ", mean(aggregate_time_in_queue_Array))
println("95% confidecn interval",ci(OneSampleTTest(aggregate_time_in_queue_Array)),"\n")

println("Mean Customers served ", mean(customers_served))
println("95% confidecn interval",ci(OneSampleTTest(customers_served)),"\n")
