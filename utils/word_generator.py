import os
import json
import random
import re
from groq import Groq
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class WordGenerator:
    def __init__(self):
        self.client = Groq(api_key=os.getenv("GROQ_API_KEY"))
        self.topics = self.load_topics()

    def load_topics(self):
        file_path = os.path.join('data', 'topics.json')
        with open(file_path, 'r') as file:        
            topics_data = json.load(file)
            topics = topics_data.get('topics', [])
            random.shuffle(topics)  # Shuffle the topics randomly
            return topics
        
    def generate_words(self, age, difficulty, num_words=10, exclude_words=[]):
        # Add randomness to the prompt
        random_seed = random.randint(1, 1000000)
        exclude_words_str = ", ".join(exclude_words)

        # Select random topics
        selected_topics = random.sample(self.topics, min(len(self.topics), 5))  # Select up to 5 topics randomly
        selected_topics_str = ", ".join(selected_topics)

        prompt = f"""
        Generate a list of {num_words} English words suitable for a {age}-year-old learning English.
        The difficulty level is {difficulty}. Use the random seed {random_seed} to ensure variety.
        The words should be from the following topics, one word per topic: {selected_topics_str}.
        DO NOT include the following words: {exclude_words_str}.
        For each word, provide:
        1. The English word
        2. Its ACCURATE Hebrew translation
        3. Three incorrect Hebrew options that are CLEARLY DIFFERENT from the correct translation
        
        It is CRUCIAL that the Hebrew translations are 100% accurate. Double-check each translation.
        If you're unsure about a translation, exclude that word and choose another.

        Format the output as a valid JSON array of objects, like this:
        [
            {{"english": "Dog", "hebrew": "כלב", "options": ["חתול", "ציפור", "דג"]}},
            {{"english": "Apple", "hebrew": "תפוח", "options": ["בננה", "תפוז", "ענב"]}},
            // More words...
        ]

        Ensure that:
        1. The words are appropriate for the age and difficulty level.
        2. The Hebrew translation is 100% correct.
        3. The incorrect options are plausible but clearly wrong and different from the correct answer.
        4. The output is a valid JSON array that can be parsed directly.
        5. None of the excluded words are included in the generated list.
        """
        print(prompt)        
        response = self.client.chat.completions.create(
            messages=[
                {
                    "role": "system",
                    "content": "You are a helpful assistant that generates word lists for language learning applications. Always respond with valid JSON."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            model=os.getenv("GROQ_MODEL"),            
            max_tokens=int(os.getenv("GROQ_MAX_TOKENS", 1000)),
            temperature=0,  # Increased temperature for more randomness            
            stop=None,
        )
        
        content = response.choices[0].message.content
        json_array_match = re.search(r'\[[\s\S]*\]', content)
        if json_array_match:
            json_array_str = json_array_match.group(0)
            try:
                words = json.loads(json_array_str)
                # Randomize the position of the correct answer for each word
                for word in words:
                    options = word['options'] + [word['hebrew']]
                    random.shuffle(options)
                    word['options'] = options
                
                 # Print the words in English as a comma-separated list
                english_words = [word['english'] for word in words]
                print(", ".join(english_words))

                return words
            except json.JSONDecodeError:
                print("Error: Invalid JSON format in the extracted array")
                return []
        else:
            print("Error: No JSON array found in the response")
            return []

if __name__ == "__main__":
    # Example usage:
    generator = WordGenerator()
    words = generator.generate_words(8, "medium", 5, ["Dog", "Cat"])
    # print(json.dumps(words, indent=2, ensure_ascii=False))

    