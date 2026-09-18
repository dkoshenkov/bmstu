module koshenkov (
	input rst,
	input clk,
	input [7:0] U,
	output [7:0] C,
);

localparam [2:0] N0 = 3'000;
localparam [2:0] N1 = 3'001;
localparam [2:0] N2 = 3'010;
localparam [2:0] N3 = 3'011;
localparam [2:0] N4 = 3'100;
localparam [2:0] N5 = 3'101;
// 76543210
// 00000000

localparam [7:0] out_zero = 8'b00000000;
// 									  76543210
localparam [7:0] out_N0_N5 = 8'b01000110;

localparam [7:0] out_N1_N4 = 8'b01111111;
localparam [7:0] out_N1_N5 = 8'b10010001;

localparam [7:0] out_N2_N2 = 8'b10000000;
localparam [7:0] out_N2_N5 = 8'b01110001;

localparam [7:0] out_N3_N0 = 8'b00111001;
localparam [7:0] out_N3_N1 = 8'b10110110;

localparam [7:0] out_N4_N3 = 8'b01111111;

localparam [7:0] out_N5_N4 = 8'b10110111;
localparam [7:0] out_N5_N1 = 8'b01111111;

reg [2:0] state;


always_ff @(posedge clk)
begin
	if rst
		begin
			state <= N0;
			C <= out_zero;
		end
	else
		begin
			case (state)
			N0:
				begin
					// N5
					if ((~U[1] & ~U[0]) | ~U[6] & U[4])
						begin
							state <= N5;
							C <= out_N1_N5;
						end
					// @
					else
						begin
							state <= N0;
							C <= out_zero;
						end
				end
			N1:
				begin
					// N4
					if (U[6] | U[2]&~U[3]&U[5]&~U[4]& U[1] | U[0])
						begin
							state <= N4;
							C <= out_N1_N4;
						end
					// N5
					else if ( U[3]&U[1] )
						begin
							state <= N5;
							C <= out_N1_N5;
						end
					// @
					else
						begin
							state <= N1;
							C <= out_zero;
						end
				end
			N2:
				begin
					// N5
					if ( U[0]&U[1] | U[2] | U[7] )
						begin
							state <= N5;
							C <= out_N2_N5;
						end
					else
					// @
						begin
							state <= N2;
							C <= out_zero;
						end
					end
			N3:
				begin
					// N0
					if (~U[2]) 
						begin
							state <= N0;
							C <= out_N3_N0;
						end
					// N1
					else if ( ~U[2]|U[1] | ~U[6]&~U[5]&~U[3] | U[7] )
						begin
							state <= N1;
							C <= out_N3_N1;
						end
					// @
					else 
						begin
							state <= N3;
							C <= out_zero;
						end
				end
			N4:
				begin
				// N0
				if (~U[0]) 
					begin
						state <= N0;
						C <= out_N4_N0;
					end
				// @
				else
					begin
						state <= N4;
						C <= out_zero;
					end
				end
			N5:
				begin
				// N1
				if (U[1] | U[6] & U[2] + ~U[7] & ~U[3] | ~U[4] & ~U[5])
					begin
						state <= N1;
						C <= out_N5_N1;
					end
				// N4
				else if (U[7])
					begin
						state <= N4;
						C <= out_N5_N4;
					end
				// @
				else 
					begin
						state <= N5;
						C <= out_zero;
					end
				end
		end
end
endmodule