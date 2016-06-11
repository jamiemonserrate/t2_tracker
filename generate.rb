#!/usr/bin/env ruby

require 'rubygems'
require 'bundler/setup'
require_relative 'lib/t2_generator'

raise "Please provide a path to a config file. Take a look at the config folder for examples." if ARGV.first.empty?

T2Generator.new(ARGV.first).run
