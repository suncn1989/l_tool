/*
 * Random select from result list.
 */

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class RandomSelect {
	
	//Generate a random number
	public List<Integer> GenRandomNum(int size)
	{
		Random ran = new Random();
		List<Integer> list = new ArrayList<Integer>();
		boolean[] bool = new boolean[size];
		int tempNum = 0;
		
		for(int i=0; i<size; i++)
		{
			do
			{
				tempNum = ran.nextInt(size);
			}
			while(bool[tempNum]);
			
			bool[tempNum] = true;
			
			list.add(tempNum);
		}
		
		return list;
	}
}
