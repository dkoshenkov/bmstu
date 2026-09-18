

module test_module(
	input clk,
	input rst,
	input cs,
	input wire [7:0] d_in,
	
	output ack,
	output err,
	output [15:0] d_out
);

logic header_written = 0;
logic b1_written = 0;
logic b2_written = 0;
logic [7:0] b1;

always @(posedge clk) 
begin
	if (rst)
	begin
		ack <= 1;
		err <= 0;
		d_out <= `z;
	end
	
	if (!cs)
	begin
		if ( !header_written )
		begin
			if (d_in == 8'b11001010)
			begin
				header_written <= 1;
				ack <= 0;
				err <= 0;
			end
			else
			begin
				ack <= 1;
				err <= 1;
			end
		end
		else if ( !b1_written )
		begin
			ack <= 1;
			err <= 0;
			b1 <= d_in;
			b1_written <= 1;
		end
		else if ( !b2_written )
		begin
			d_out <= {b1, d_in};
			b2_written<=1;
		end
		else
		begin
			d_out <= 'z;
			err <=0;
			ack <=1;
			b1_written <= 0;
			b2_written <= 0;
			header_written <= 0;
		end
	end
end



endmodule

