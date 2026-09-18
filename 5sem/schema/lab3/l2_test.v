`timescale 1ns / 1ps

////////////////////////////////////////////////////////////////////////////////
// Company: 
// Engineer:
//
// Create Date:   15:59:20 06/07/2024
// Design Name:   l2
// Module Name:   /home/user/lab2_o_gore/l2_test.v
// Project Name:  lab2_o_gore
// Target Device:  
// Tool versions:  
// Description: 
//
// Verilog Test Fixture created by ISE for module: l2
//
// Dependencies:
// 
// Revision:
// Revision 0.01 - File Created
// Additional Comments:
// 
////////////////////////////////////////////////////////////////////////////////

module l2_test;

	// Outputs
	reg rst;
	reg clk;
	reg count;

	// Outputs
	wire cnt;
	wire [1:0] s_out;

	// Instantiate the Unit Under Test (UUT)
	l2 uut (
		.rst(rst), 
		.clk(clk), 
		.count(count), 
		.cnt(cnt), 
		.s_out(s_out)
	);

	always #10 clk=!clk;
	initial begin
		// Initialize Inputs
		rst=1;
		clk=0;
		count=0;
		#600
		rst=0;
		count=1;
		
		#100
		count=0;
	
		#100
		count=1;
		// Wait 100 ns for global reset to finish
		#100
		count=0;
		
      #100
		count=1;
		#100
		count=0;
	
		#100
		count=1;
		// Wait 100 ns for global reset to finish
		#100
		count=0;
		
		#100
		count=1;
		
		#4000
		count=0;
		
		#100
		count=1;
		#100
		count=0;
	
		#100
		count=1;
		// Wait 100 ns for global reset to finish
		#100
		count=0;
		
		#100
		count=1;
		#100
		count=0;
	
		#100
		count=1;
		// Wait 100 ns for global reset to finish
		#100
		count=0;
		
		#100
		count=1;
		#100
		count=0;
	
		#100
		count=1;
		// Wait 100 ns for global reset to finish
		#100
		count=0;
		
		#100
		count=1;
		#100
		count=0;
	
		#100
		count=1;
		// Wait 100 ns for global reset to finish
		#100
		count=0;
		

	end

      
endmodule

