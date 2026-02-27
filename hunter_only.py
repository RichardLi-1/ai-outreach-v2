import openai_hunter_client

def main():
    while True:
        firstName = input("First Name: ")
        lastName = input("Last Name: ")
        domain = input("Domain: ")
        res = openai_hunter_client.find_email(firstName, lastName, domain)
        if str(res[0]) == "200":
            parsedEmailResponse = res[1].get("data")
            print(parsedEmailResponse)
            email = parsedEmailResponse.get("email")
            if email is not None:
                score = parsedEmailResponse["score"]
                sources = parsedEmailResponse.get("sources", [])
                source = sources[0]["uri"] if sources else ""
                print("Data found by hunter.io:")
                print("email: " + email)
                print("score: " + str(score))
                print("source: " + source)
            else:
                print("Hunter.io did not find an email for " + firstName + " " + lastName)
        else:
            print("Hunter.io failed to find")

        


if __name__ == "__main__":
    print("Directly interface with hunter.io API")
    main()