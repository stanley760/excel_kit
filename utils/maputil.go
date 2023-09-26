package utils

// ContainsKey compare with the special key.
func ContainsKey(m map[string][]string, key string) bool {
	for k := range m {
		if k == key {
			return true
		}
	}
	return false
}
