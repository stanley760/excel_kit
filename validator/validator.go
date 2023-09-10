package validator

import (
	"fmt"
	"github.com/fzdwx/infinite/components"
)

func Min(numberItems int, temp string) components.Validator {
	// return a validator that checks the length of the List
	return func(val interface{}) error {
		if list, ok := val.([]int); ok {
			// if the List is shorter than the given value
			if len(list) < numberItems {
				// yell loudly
				return fmt.Errorf(temp, numberItems)

			}
		} else {
			// otherwise we cannot convert the value into a List of answer and cannot enforce length
			return fmt.Errorf("无法读取excel工作薄")
		}
		// the input is fine
		return nil
	}
}
