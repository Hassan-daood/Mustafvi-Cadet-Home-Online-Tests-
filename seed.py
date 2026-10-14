from app import app, db, Test, Section, Question

with app.app_context():
    if Section.query.count() == 0:
        test = Test(title="Demo Test")
        db.session.add(test)
        db.session.flush()
        s1 = Section(test_id=test.id, title="Basic Knowledge", duration_minutes=5, order_no=1)
        s2 = Section(test_id=test.id, title="Second Section", duration_minutes=5, order_no=2)
        db.session.add_all([s1, s2])
        db.session.flush()
        db.session.add_all([
            Question(section_id=s1.id, text="What is 2 + 2?", option_a="3", option_b="4", option_c="5", option_d="6", correct_option="B"),
            Question(section_id=s1.id, text="Which colour is the sky on a clear day?", option_a="Blue", option_b="Green", option_c="Red", option_d="Black", correct_option="A"),
            Question(section_id=s2.id, text="What is 5 x 3?", option_a="8", option_b="10", option_c="15", option_d="20", correct_option="C"),
            Question(section_id=s2.id, text="Which one is a programming language?", option_a="Python", option_b="Chair", option_c="Window", option_d="River", correct_option="A"),
        ])
        db.session.commit()
        print("Demo sections and questions added.")
    else:
        print("Database already contains sections; nothing added.")
