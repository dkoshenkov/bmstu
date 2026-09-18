module rk(
	input clk,
	input rst,
	input cs,
	input [7:0] d_in,
	
	output d_out1,
	output d_out2,
	output valid
);

logic [7:0] D;

always @(posedge clk)
begin
	if (!rst) 
	begin
		d_out1 <= 'x;
		d_out2 <= 'x;
		valid <= 1;
	end
	else if (!сs)
	begin
		
	end	
end

endmodule
