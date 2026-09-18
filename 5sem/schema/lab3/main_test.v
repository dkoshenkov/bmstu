`timescale 1ns / 1ps

////////////////////////////////////////////////////////////////////////////////
// Company: 
// Engineer:
//
// Create Date:   17:00:42 06/07/2024
// Design Name:   main
// Module Name:   /home/user/lab2_o_gore/main_test.v
// Project Name:  lab2_o_gore
// Target Device:  
// Tool versions:  
// Description: 
//
// Verilog Test Fixture created by ISE for module: main
//
// Dependencies:
// 
// Revision:
// Revision 0.01 - File Created
// Additional Comments:
// 
////////////////////////////////////////////////////////////////////////////////

module main_test;

	// Inputs
	reg clk;
	reg count;
	reg rst;

	// Outputs
	wire [3:0] a;
	wire [7:0] led;
	wire [1:0] state;

	// Instantiate the Unit Under Test (UUT)
	main uut (
		.clk(clk), 
		.count(count), 
		.rst(rst), 
		.a(a), 
		.led(led), 
		.state(state)
	);

	initial begin
		// Initialize Inputs
		clk = 0;
		count = 0;
		rst = 0;

		// Wait 100 ns for global reset to finish
		#100;
        
		// Add stimulus here

	end
      
endmodule

